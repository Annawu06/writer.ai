import uno
import re
import unohelper
import json
import os
import traceback
import sys # To print to stderr
import threading
from collections import deque
from format_plan import (
    parse_model_json, validate_format_request, count_format_operations,
    changes_document_content,
)
from urllib.error import HTTPError, URLError
from urllib.request import Request, urlopen
from urllib.parse import urlparse
from com.sun.star.task import XJobExecutor
from com.sun.star.util import XModifyListener
from com.sun.star.awt import XActionListener, XItemListener
from com.sun.star.awt import XCallback
from com.sun.star.awt.PosSize import SIZE, POSSIZE
from com.sun.star.awt.PushButtonType import OK, CANCEL

from com.sun.star.awt.FontWeight import BOLD
from com.sun.star.awt.FontSlant import ITALIC
from com.sun.star.awt import FontUnderline
from com.sun.star.style.ParagraphAdjust import LEFT, RIGHT, CENTER, BLOCK

from com.sun.star.ui.dialogs.TemplateDescription import FILEOPEN_SIMPLE

from com.sun.star.awt.MessageBoxButtons import BUTTONS_OK, BUTTONS_YES_NO
from com.sun.star.awt.MessageBoxResults import YES

from com.sun.star.awt.FontWeight import NORMAL
from com.sun.star.awt.FontSlant import NONE
from com.sun.star.awt.FontUnderline import NONE as UNDERLINE_NONE
from com.sun.star.beans.PropertyState import AMBIGUOUS_VALUE


# select file to format
def pick_writer_file(ctx):

    smgr = ctx.getServiceManager()

    file_picker = smgr.createInstanceWithContext(
        "com.sun.star.ui.dialogs.FilePicker",
        ctx
    )

    file_picker.initialize((FILEOPEN_SIMPLE,))

    file_picker.setTitle("Select Writer Document")

    file_picker.appendFilter("Writer Documents", "*.odt;*.docx;*.doc")
    file_picker.setCurrentFilter("Writer Documents")

    result = file_picker.execute()

    if result == 1:
        files = file_picker.getFiles()
        return files[0]

    return None
    
# Helper for debugging
def log_to_console(*args):
    """Prints messages to the console for debugging."""
    print(*args, file=sys.stderr)
    sys.stderr.flush()


_FORMAT_REQUEST_QUEUE = deque()
_FORMAT_REQUEST_QUEUE_LOCK = threading.Lock()
_FORMAT_REQUEST_ACTIVE = False


class DocumentGuard(unohelper.Base, XModifyListener):
    """Detect edits and selection movement while a model request is pending."""

    def __init__(self, doc):
        self.doc = doc
        self.changed = False
        self.disposed = False
        self.selection = None
        try:
            selected = doc.getCurrentController().getSelection()
            if selected and selected.getCount() == 1:
                self.selection = doc.Text.createTextCursorByRange(selected.getByIndex(0))
        except Exception:
            pass
        doc.addModifyListener(self)

    def modified(self, event):
        self.changed = True

    def disposing(self, event):
        self.disposed = True

    def selection_matches(self):
        if self.selection is None:
            return False
        try:
            selected = self.doc.getCurrentController().getSelection()
            if not selected or selected.getCount() != 1:
                return False
            current = selected.getByIndex(0)
            text = self.doc.Text
            return (text.compareRegionStarts(self.selection, current) == 0
                    and text.compareRegionEnds(self.selection, current) == 0)
        except Exception:
            return False

    def can_apply(self, plan):
        if self.changed or self.disposed:
            return False
        try:
            if self.doc.getCurrentController() is None:
                return False
        except Exception:
            return False
        return "selection" not in plan or self.selection_matches()

    def close(self):
        if not self.disposed:
            try:
                self.doc.removeModifyListener(self)
            except Exception:
                pass


def _run_next_format_request():
    global _FORMAT_REQUEST_ACTIVE
    with _FORMAT_REQUEST_QUEUE_LOCK:
        if _FORMAT_REQUEST_ACTIVE or not _FORMAT_REQUEST_QUEUE:
            return
        job, target_doc, query, api_key, model, endpoint, guard = _FORMAT_REQUEST_QUEUE.popleft()
        _FORMAT_REQUEST_ACTIVE = True
    try:
        job._launch_format_request(target_doc, query, api_key, model, endpoint, guard)
    except Exception as error:
        log_to_console(f"Unable to start queued formatting request: {error}")
        guard.close()
        _finish_queued_format_request()


def _finish_queued_format_request():
    global _FORMAT_REQUEST_ACTIVE
    with _FORMAT_REQUEST_QUEUE_LOCK:
        _FORMAT_REQUEST_ACTIVE = False
    _run_next_format_request()


def get_doc(ctx):
    smgr = ctx.getServiceManager()
    desktop = smgr.createInstanceWithContext(
        "com.sun.star.frame.Desktop", ctx
    )
    return desktop.getCurrentComponent()


def describe_plan_targets(plan, fmt):
    """Describe the planned target range without modifying the document."""
    descriptions = []
    for scope, options in plan.items():
        if scope in {"all_pages", "document", "entire_doc"}:
            descriptions.append(f"{scope}: entire document")
        elif scope == "selection":
            descriptions.append("selection: current selection")
        elif re.fullmatch(r"paragraph_\d+", scope) or scope == "paragraph_text":
            descriptions.append(f"{scope}: 1 paragraph")
        elif scope == "paragraphs":
            descriptions.append(f"paragraphs: {len(options)} targets")
        elif scope == "structure":
            counts = {role: 0 for role in options if role != "infer"}
            first_content_seen = False
            plain_roles = {
                "Title": "title", "Subtitle": "title", "Heading 1": "heading_1",
                "Heading 2": "heading_2", "Heading 3": "heading_3",
                "Text Body": "body", "Body Text": "body",
                "Default Paragraph Style": "body", "Standard": "body",
            }
            for paragraph in fmt._paragraphs():
                if not paragraph.String.strip():
                    continue
                role = (fmt._paragraph_role(paragraph, first_content_seen)
                        if options.get("infer") else plain_roles.get(paragraph.ParaStyleName))
                if role in counts:
                    counts[role] += 1
                first_content_seen = True
            descriptions.extend(f"structure.{role}: {count} paragraph{'s' if count != 1 else ''}"
                                for role, count in counts.items())
        elif scope == "tables":
            for table_scope in options:
                count = len(fmt.get_tables(table_scope))
                descriptions.append(f"tables.{table_scope}: {count} table{'s' if count != 1 else ''}")
        elif re.fullmatch(r"table_\d+", scope):
            count = len(fmt.get_tables(scope))
            descriptions.append(f"{scope}: {count} table{'s' if count != 1 else ''}")
    return descriptions


def _paragraph_property(paragraph, name):
    """Return a visible value, including mixed character formatting."""
    try:
        if paragraph.getPropertyState(name) == AMBIGUOUS_VALUE:
            return "mixed"
        return getattr(paragraph, name)
    except Exception:
        return "unknown"


def _paragraph_changes(paragraph, style, implied_style=None):
    """Compare supported paragraph properties with the requested values."""
    changes = []
    requested_style = implied_style or style.get("paragraph_style")
    if style.get("title"):
        requested_style = "Title"
    if style.get("body"):
        requested_style = "Text Body"
    if requested_style:
        current = _paragraph_property(paragraph, "ParaStyleName")
        if current != requested_style:
            changes.append(f"style: {current} → {requested_style}")

    for key, property_name, display in (
        ("bold", "CharWeight", lambda value: "on" if value >= BOLD else "off"),
        ("font_size", "CharHeight", lambda value: f"{value:g} pt"),
        ("font_color", "CharColor", lambda value: "automatic" if value < 0 else f"#{value:06X}"),
    ):
        if key not in style:
            continue
        current = _paragraph_property(paragraph, property_name)
        try:
            current_text = display(current) if current not in {"mixed", "unknown"} else current
        except (TypeError, ValueError):
            current_text = "unknown"
        requested = style[key]
        if key == "bold":
            requested = "on" if requested else "off"
        elif key == "font_size":
            requested = f"{requested:g} pt"
        elif isinstance(requested, int):
            requested = f"#{requested:06X}"
        else:
            requested = "#" + requested.lstrip("#").upper()
        if current_text != requested:
            changes.append(f"{key}: {current_text} → {requested}")

    alignments = {
        "align_left": (LEFT, "left"), "align_right": (RIGHT, "right"),
        "align_center": (CENTER, "center"), "align_justify": (BLOCK, "justify"),
    }
    current_alignment = _paragraph_property(paragraph, "ParaAdjust")
    current_alignment = next(
        (name for value, name in alignments.values() if current_alignment == value),
        current_alignment if current_alignment in {"mixed", "unknown"} else "unknown",
    )
    for key, (_, requested) in alignments.items():
        if style.get(key) and current_alignment != requested:
            changes.append(f"alignment: {current_alignment} → {requested}")

    if "first_line_indent" in style:
        current = _paragraph_property(paragraph, "ParaFirstLineIndent")
        current_text = f"{current / 100:.1f} mm" if isinstance(current, (int, float)) else current
        changes.append(f"first-line indent: {current_text} → {style['first_line_indent']} characters")
    return changes


def _table_changes(table, options, fmt):
    """Compare table pagination and the first 20 affected cell backgrounds."""
    changes = []
    for key, property_name in (("split", "Split"), ("keep_together", "KeepTogether")):
        if key in options:
            current = bool(getattr(table, property_name))
            if current != options[key]:
                changes.append(f"{key}: {current} → {options[key]}")
    if options.get("repeat_header") and not table.RepeatHeadline:
        changes.append("repeat_header: False → True")

    color_keys = {"table_background", "cell_background", "header_background", "zebra", "zebra_background"}
    if not color_keys.intersection(options):
        return changes
    header = options.get("header", options.get("format_header", False))
    names = table.getCellNames()
    for cell_name in names[:20]:
        row = fmt._cell_row(cell_name)
        cell_options = dict(options)
        if options.get("zebra") and row and row % 2 == 0:
            cell_options["cell_background"] = options.get("zebra_background", "F2F2F2")
        color = cell_options.get("header_background" if header and row == 1 else "cell_background")
        if color is None:
            color = cell_options.get("table_background")
        if color is None:
            continue
        cell = table.getCellByName(cell_name)
        current = "transparent" if cell.BackTransparent else f"#{cell.BackColor:06X}"
        requested = f"#{fmt.parse_color(color):06X}"
        if current != requested:
            changes.append(f"{cell_name} background: {current} → {requested}")
    if len(names) > 20:
        changes.append(f"{len(names) - 20} additional cells omitted from comparison")
    return changes


def describe_plan_differences(plan, fmt):
    """Preview comparable properties; leave other operations in the full plan."""
    paragraphs = list(fmt._paragraphs())
    entries = []

    def add(label, paragraph, style, implied_style=None):
        changes = _paragraph_changes(paragraph, style, implied_style)
        if changes:
            entries.append(f"{label}: " + "; ".join(changes))

    for scope, options in plan.items():
        if scope == "selection":
            cursor = fmt.get_selection_cursor()
            if cursor is not None:
                add("selection", cursor, options)
        elif re.fullmatch(r"paragraph_\d+", scope):
            index = int(scope.split("_")[1]) - 1
            add(scope, paragraphs[index], options)
        elif scope == "paragraph_text":
            needle = options["text"].casefold()
            match = next((p for p in paragraphs if needle in p.String.casefold()), None)
            if match is not None:
                add("paragraph_text", match, options["format"])
        elif scope == "paragraphs":
            for target, style in options.items():
                if re.fullmatch(r"paragraph_\d+", target):
                    add(target, paragraphs[int(target.split("_")[1]) - 1], style)
                else:
                    needle = style["text"].casefold()
                    match = next((p for p in paragraphs if needle in p.String.casefold()), None)
                    if match is not None:
                        add("paragraphs.contains", match, style["format"])
        elif scope == "structure":
            plain_roles = {
                "Title": "title", "Subtitle": "title", "Heading 1": "heading_1",
                "Heading 2": "heading_2", "Heading 3": "heading_3",
                "Text Body": "body", "Body Text": "body",
                "Default Paragraph Style": "body", "Standard": "body",
            }
            style_names = {
                "title": "Title", "heading_1": "Heading 1", "heading_2": "Heading 2",
                "heading_3": "Heading 3", "body": "Text Body",
            }
            first_content_seen = False
            for index, paragraph in enumerate(paragraphs, start=1):
                if not paragraph.String.strip():
                    continue
                role = (fmt._paragraph_role(paragraph, first_content_seen) if options.get("infer")
                        else plain_roles.get(paragraph.ParaStyleName))
                if role in options:
                    style = options[role] if isinstance(options[role], dict) else {}
                    add(f"paragraph_{index} ({role})", paragraph, style, style_names[role])
                first_content_seen = True
        elif scope == "tables" or re.fullmatch(r"table_\d+", scope):
            table_options = options if scope == "tables" else {scope: options}
            for table_scope, style in table_options.items():
                for table in fmt.get_tables(table_scope):
                    changes = _table_changes(table, style, fmt)
                    if changes:
                        entries.append(f"{table_scope}: " + "; ".join(changes))
    if not entries:
        return ["No comparable property changes detected; review the full plan for other operations."]
    if len(entries) > 50:
        return entries[:50] + [f"… {len(entries) - 50} more paragraph changes; review the full plan."]
    return entries


class Format:

    def __init__(self, ctx, doc):

        self.ctx = ctx
        self.doc = doc

        if self.doc is None:
            raise RuntimeError("No active document")

        self.controller = self.doc.getCurrentController()
        
        

    def parse_color(self, color):
        # R*65536 + G*256 + B
        if isinstance(color, int):
            return color
        std_colors = {
            "white": 0xFFFFFF, "black": 0x000000, "gray": 0x808080, "silver": 0xC0C0C0,
            "darkgray": 0xA9A9A9,"red": 0xFF0000, "darkred": 0x8B0000, "maroon": 0x800000,
            "orange": 0xFFA500, "yellow": 0xFFFF00,
            "olive": 0x808000, "lime": 0x00FF00, "green": 0x008000, "aqua": 0x00FFFF,
            "teal": 0x008080, "blue": 0x0000FF, "navy": 0x000080, "fuchsia": 0xFF00FF,
            "purple": 0x800080, "pink": 0xFFC0CB, "gold": 0xFFD700, "brown": 0xA52A2A,
            "cyan": 0x00FFFF, "magenta": 0xFF00FF,"tiffanyblue": 0x0ABAB5
        }

        # process highlight is true
        if color is True:
            return std_colors["yellow"]
        
        if not color or not isinstance(color, str):
            return std_colors["yellow"]

        clean_color = color.lower().replace(" ", "").strip().lstrip('#')

        if clean_color in std_colors:
            return std_colors[clean_color]

        if re.fullmatch(r'[0-9a-f]{3}|[0-9a-f]{6}', clean_color):
            try:
                if len(clean_color) == 3:
                    clean_color = ''.join([c*2 for c in clean_color])
                return int(clean_color, 16)
            except ValueError:
                pass

        return std_colors["yellow"]
        
    def get_cursor(self):

        return self.controller.getViewCursor()
        
    def get_document_cursor(self):
        """obtain TextCursor"""
        cursor = self.doc.Text.createTextCursor()
        cursor.gotoStart(False) 
        cursor.gotoEnd(True)  
        return cursor

    def _paragraphs(self):
        paragraphs = self.doc.Text.createEnumeration()
        while paragraphs.hasMoreElements():
            element = paragraphs.nextElement()
            if element.supportsService("com.sun.star.text.Paragraph"):
                yield element

    def get_paragraph_cursor(self, paragraph_number):
        """Return a cursor for a document paragraph using a stable ordinal."""
        if paragraph_number < 1:
            return None
        for index, paragraph in enumerate(self._paragraphs(), start=1):
            if index == paragraph_number:
                cursor = self.doc.Text.createTextCursorByRange(paragraph.getStart())
                cursor.gotoRange(paragraph.getEnd(), True)
                return cursor
        return None

    def find_paragraph_cursor(self, text):
        """Resolve an unambiguous paragraph match, independent of pagination."""
        needle = str(text).casefold()
        if not needle:
            raise ValueError("Paragraph search text is empty")
        matches = [paragraph for paragraph in self._paragraphs()
                   if needle in paragraph.String.casefold()]
        if len(matches) != 1:
            raise ValueError(f"Paragraph search must match exactly one paragraph; found {len(matches)}")
        cursor = self.doc.Text.createTextCursorByRange(matches[0].getStart())
        cursor.gotoRange(matches[0].getEnd(), True)
        return cursor

    @staticmethod
    def _is_probable_heading(text):
        text = text.strip()
        return bool(text) and len(text) <= 40 and not re.search(r"[。！？.!?；;]$", text)

    @staticmethod
    def _heading_level_from_style(style):
        match = re.fullmatch(r"Heading ([1-3])", str(style))
        return int(match.group(1)) if match else None

    @staticmethod
    def _inferred_heading_level(text):
        text = text.strip()
        if re.match(r"^\d+\.\d+(?:\.\d+)?[、. ]+", text):
            return 2
        if re.match(r"^(?:第[一二三四五六七八九十百千万]+章|\d+[、. )])\s*", text):
            return 1
        return None

    def _paragraph_role(self, paragraph, first_content_seen):
        style = paragraph.ParaStyleName
        if style in {"Title", "Subtitle"}:
            return "title"
        if not first_content_seen and style in {
            "Text Body", "Body Text", "Default Paragraph Style", "Standard"
        }:
            return "title"
        heading_level = self._heading_level_from_style(style)
        if heading_level:
            return f"heading_{heading_level}"
        inferred_level = self._inferred_heading_level(paragraph.String)
        if inferred_level:
            return f"heading_{inferred_level}"
        if style in {"Text Body", "Body Text", "Default Paragraph Style", "Standard"}:
            return "body"
        if not first_content_seen:
            return "title"
        if self._is_probable_heading(paragraph.String):
            return "heading_1"
        return "body"

    def format_document_structure(self, options):
        """Apply semantic paragraph roles without touching paragraphs in tables."""
        if not isinstance(options, dict):
            return 0
        infer = options.get("infer", False) is True
        first_content_seen = False
        roles = {"title", "heading_1", "heading_2", "heading_3", "body"}
        paragraphs = list(self._paragraphs())
        style_assignments = []
        indent_assignments = []
        matched_count = 0
        for paragraph_index, paragraph in enumerate(paragraphs):
            if not paragraph.String.strip():
                continue
            role = self._paragraph_role(paragraph, first_content_seen) if infer else {
                "Title": "title", "Subtitle": "title", "Heading 1": "heading_1",
                "Heading 2": "heading_2", "Heading 3": "heading_3",
                "Text Body": "body", "Body Text": "body",
                "Default Paragraph Style": "body", "Standard": "body",
            }.get(paragraph.ParaStyleName)
            if role not in roles or role not in options:
                first_content_seen = True
                continue

            cursor = self.doc.Text.createTextCursorByRange(paragraph.getStart())
            cursor.gotoRange(paragraph.getEnd(), True)
            role_options = options[role]
            if role_options is True:
                role_options = {}
            if not isinstance(role_options, dict):
                first_content_seen = True
                continue
            matched_count += 1
            style_name = {
                "title": "Title", "heading_1": "Heading 1",
                "heading_2": "Heading 2", "heading_3": "Heading 3",
                "body": "Text Body",
            }[role]
            apply_styles(self, cursor, role_options)
            if "first_line_indent" in role_options:
                indent_assignments.append((paragraph_index, role_options["first_line_indent"]))
            available_style = self.available_paragraph_style(style_name)
            if available_style:
                style_assignments.append((paragraph_index, available_style))
            first_content_seen = True

        refreshed_paragraphs = list(self._paragraphs())
        for paragraph_index, style_name in style_assignments:
            if paragraph_index < len(refreshed_paragraphs):
                refreshed_paragraphs[paragraph_index].ParaStyleName = style_name
        for paragraph_index, character_count in indent_assignments:
            if paragraph_index < len(refreshed_paragraphs):
                try:
                    paragraph = refreshed_paragraphs[paragraph_index]
                    font_size = float(getattr(paragraph, "CharHeight", 12) or 12)
                    paragraph.ParaFirstLineIndent = int(
                        float(character_count) * font_size * 2540 / 72
                    )
                except (TypeError, ValueError) as e:
                    log_to_console(f"Invalid structural first-line indent: {e}")
        return matched_count
        
    def get_selection(self):
        """
                获取选中的文本对象
        """
        selection = self.controller.getSelection()
        if selection.getCount() > 0:
            return selection.getByIndex(0)
        return None
        
    # ------------------------------------------------
    # Text Style
    # ------------------------------------------------

    def set_bold(self,cursor, value=True):
        cursor.CharWeight = BOLD if value is not False else NORMAL

    def set_italic(self,cursor, value=True):
        cursor.CharPosture = ITALIC if value is not False else NONE

    def set_underline(self, cursor, value):
        try:
            if value is False:
                cursor.CharUnderline = 0
                cursor.CharUnderlineHasColor = False
                return
            val_str = str(value).strip()
            style_part = "1"
            color_part = None

  
            if len(val_str) >= 7:
                style_part = val_str[:-6]
                color_part = val_str[-6:]
            elif len(val_str) > 0:
                style_part = val_str
                color_part = None
            
            try:
                style_int = int(style_part)
                cursor.CharUnderline = style_int if 0 <= style_int <= 18 else 1
            except ValueError:
                cursor.CharUnderline = 1

            if color_part:
                cursor.CharUnderlineHasColor = True
                cursor.CharUnderlineColor = int(color_part, 16)
            else:
                cursor.CharUnderlineHasColor = False

        except Exception as e:
            raise RuntimeError("Unable to set underline") from e


    def set_font_name(self, cursor, font_name):
            try:
                if not font_name or not isinstance(font_name, str):
                    return

                font_map = {
                    # --- 基础类别 ---
                    "serif": "Libre Serif",
                    "sans-serif": "Libre Sans",
                    "monospace": "Liberation Mono",
                    "code": "Consolas",
                    
                    "modern": "Noto Sans",
                    "clean": "DejaVu Sans",
                    "minimal": "Inter",
                    
                    "formal": "Libre Baskerville",
                    "academic": "Linux Libertine G",
                    "professional": "Liberation Serif",
                    "classic": "Times New Roman",
                    
                    "chinese": "Noto Sans CJK SC",
                    "heiti": "Noto Sans CJK SC",
                    "songti": "Noto Serif CJK SC",
                    "kaiti": "AR PL UKai CN",
                    "microsoft yahei": "Microsoft YaHei",
                    
                    "handwriting": "Comic Sans MS", 
                    "elegant": "Apple Chancery",
                    "title": "Linux Biolinum G"
                }
                

                clean_name = font_name.lower().replace(" ", "").replace("-", "")
                target_font = font_map.get(clean_name, font_name)
                
                cursor.CharFontName = target_font          
                cursor.CharFontNameAsian = target_font     
                cursor.CharFontNameComplex = target_font   
                
            except Exception as e:
                raise RuntimeError("Unable to set font name") from e



    def set_font_size(self, cursor, size):
        size = float(size)

        cursor.CharHeight = size
        cursor.CharHeightAsian = size


    def set_font_color(self, cursor, rgb):
        try:
            if isinstance(rgb, str):
                rgb = self.parse_color(rgb)
            
            cursor.CharColor = int(rgb) 
        except Exception as e:
            raise RuntimeError("Unable to set font color") from e


    def highlight(self, cursor, color=None):

        if color is None or color is True:
            color = "yellow"

        uno_color = self.parse_color(color)

        cursor.CharBackColor = uno_color
            

    def remove_highlight(self,cursor):
        cursor = cursor
        cursor.CharBackColor = -1


    # ------------------------------------------------
    # Paragraph Align
    # ------------------------------------------------

    def align_center(self, cursor):
        cursor.ParaAdjust = CENTER

    def align_left(self, cursor):
        cursor.ParaAdjust = LEFT

    def align_right(self, cursor):
        cursor.ParaAdjust = RIGHT

    def align_justify(self, cursor):
        cursor.ParaAdjust = BLOCK

    # ------------------------------------------------
    # Text insert and replace
    # ------------------------------------------------

    def get_selection_cursor(self):
            """
            获取当前文档中用户实际选中的范围。
            """
            try:
                # 修复点：将 self.target_doc 改为 self.doc
                selection = self.doc.getCurrentController().getSelection()
                if selection and selection.getCount() > 0:
                    selected_range = selection.getByIndex(0)
                    if selected_range.getString():
                        return self.doc.Text.createTextCursorByRange(selected_range)
            except Exception as e:
                log_to_console(f"Error getting selection cursor: {e}")
            return None

    def insert_text_at_cursor(self, cursor, text, insert_before=True):
        """
        在光标当前位置插入文本。
        """
        try:
            # 获取文本对象（XText）
            text_obj = cursor.getText()
            
            if insert_before:
                # 折叠到开始位置
                cursor.collapseToStart()
            else:
                # 折叠到结束位置
                cursor.collapseToEnd()
            
            # 使用 insertString 明确执行“插入”动作
            # 参数2: 要插入的字符串
            # 参数3: 是否替换当前选区 (False 表示不替换，即插入)
            text_obj.insertString(cursor, text, False)
            
            # 插入后，建议再次 collapseToEnd，防止后续指令误伤新插入的文本
            cursor.collapseToEnd()
            
        except Exception as e:
            raise RuntimeError("Unable to insert text") from e


    def replace_selection(self, cursor, text):
        try:
            if cursor is None:
                raise ValueError("No target cursor for replacement")
            cursor.setString(text)
        except Exception as e:
            raise RuntimeError("Unable to replace text") from e

    # 确保这个函数和 replace_selection 对齐
    def get_selected_text(self, cursor):
        if cursor:
            return cursor.getString()
        return ""

    def clear_format(self, cursor):

        cursor.CharWeight = NORMAL
        cursor.CharPosture = NONE
        cursor.CharUnderline = UNDERLINE_NONE

        cursor.CharStrikeout = 0

        cursor.CharColor = -1
        cursor.CharBackColor = -1

        cursor.CharHeight = 12

    # ------------------------------------------------
    # Paragraph styles and tables
    # ------------------------------------------------

    def set_paragraph_style(self, cursor, style_name):
        """Apply a Writer paragraph style to all paragraphs in the cursor."""
        if cursor and style_name:
            styles = self.doc.StyleFamilies.getByName("ParagraphStyles")
            if not styles.hasByName(style_name):
                if style_name in {"Text Body", "Body Text"} and styles.hasByName("Standard"):
                    style_name = "Standard"
                else:
                    log_to_console(f"Paragraph style is unavailable: {style_name}")
                    return
            cursor.ParaStyleName = style_name

    def available_paragraph_style(self, style_name):
        styles = self.doc.StyleFamilies.getByName("ParagraphStyles")
        if styles.hasByName(style_name):
            return style_name
        if style_name in {"Text Body", "Body Text"} and styles.hasByName("Standard"):
            return "Standard"
        return None

    def format_title(self, cursor, value=True):
        if value is not False:
            self.set_paragraph_style(cursor, "Title")

    def format_body(self, cursor, value=True):
        if value is not False:
            self.set_paragraph_style(cursor, "Text Body")

    def set_first_line_indent(self, cursor, characters):
        """Set first-line indentation; Writer stores paragraph distances in 1/100 mm."""
        try:
            char_count = float(characters)
            if char_count < 0:
                return
            font_size = float(getattr(cursor, "CharHeight", 12) or 12)
            cursor.ParaFirstLineIndent = int(char_count * font_size * 2540 / 72)
        except (TypeError, ValueError) as e:
            raise RuntimeError("Invalid first-line indent") from e

    def _table_names(self):
        return list(self.doc.TextTables.getElementNames())

    def get_tables(self, table_key):
        """Resolve table_1/table_all or a native Writer table name."""
        names = self._table_names()
        if table_key in ("table_all", "all"):
            return [self.doc.TextTables.getByName(name) for name in names]
        if table_key in names:
            return [self.doc.TextTables.getByName(table_key)]

        match = re.fullmatch(r"table_(\d+)", str(table_key))
        if match:
            index = int(match.group(1)) - 1
            if 0 <= index < len(names):
                return [self.doc.TextTables.getByName(names[index])]
        return []

    @staticmethod
    def _cell_row(cell_name):
        match = re.search(r"(\d+)$", cell_name)
        return int(match.group(1)) if match else None

    @staticmethod
    def _column_name(column_number):
        name = ""
        while column_number > 0:
            column_number, remainder = divmod(column_number - 1, 26)
            name = chr(65 + remainder) + name
        return name

    def get_table_cell_cursor(self, table, row, column):
        cell_name = f"{self._column_name(column)}{row}"
        if cell_name not in table.getCellNames():
            return None
        cell = table.getCellByName(cell_name)
        cursor = cell.createTextCursor()
        cursor.gotoEnd(True)
        return cursor

    def _insert_table_caption(self, table, options):
        caption = options.get("caption", options.get("title"))
        if not isinstance(caption, str) or not caption.strip():
            return
        try:
            table_name = getattr(table, "Name", "")
            names = self._table_names()
            table_number = names.index(table_name) + 1 if table_name in names else 1
            number = options.get("number", table_number)
            prefix = str(options.get("caption_prefix", "表")).strip() or "表"
            caption_text = f"{prefix}{number}：{caption.strip()}\n"
            if f"{prefix}{number}：" in self.doc.Text.String:
                return

            anchor = table.getAnchor()
            try:
                cursor = self.doc.Text.createTextCursorByRange(anchor)
                cursor.collapseToStart()
            except Exception:
                # A table in the document's only empty paragraph has no usable
                # outer text range; the document cursor still inserts before it.
                cursor = self.doc.Text.createTextCursor()
                cursor.gotoStart(False)
            self.doc.Text.insertString(cursor, caption_text, False)
            cursor.gotoStartOfParagraph(False)
            cursor.gotoEndOfParagraph(True)
            style_name = self.available_paragraph_style("Caption")
            if style_name:
                cursor.ParaStyleName = style_name
            cursor.ParaAdjust = CENTER
        except Exception as e:
            raise RuntimeError("Unable to insert table caption") from e

    def _format_table_cell(self, cell, options, is_header, is_first_column=False):
        cell_cursor = cell.createTextCursor()
        cell_cursor.gotoEnd(True)

        background = options.get("header_background" if is_header else "cell_background")
        if background is None:
            background = options.get("table_background")
        if background is not None:
            cell.BackColor = self.parse_color(background)
            cell.BackTransparent = False

        if is_header and options.get("header_bold", True):
            cell_cursor.CharWeight = BOLD
        if "font_name" in options or "font_family" in options:
            self.set_font_name(cell_cursor, options.get("font_name", options.get("font_family")))
        if "font_size" in options:
            self.set_font_size(cell_cursor, options["font_size"])
        if "font_color" in options:
            self.set_font_color(cell_cursor, options["font_color"])

        alignment = str(options.get("align", "")).lower()
        alignments = {"left": LEFT, "right": RIGHT, "center": CENTER, "justify": BLOCK}
        if options.get("auto_align", False):
            cell_text = cell.String.strip()
            if re.fullmatch(r"[+-]?(?:\d+(?:\.\d+)?|\d{1,3}(?:,\d{3})+)%?", cell_text):
                cell_cursor.ParaAdjust = RIGHT
            elif re.fullmatch(r"\d{4}[-/]\d{1,2}[-/]\d{1,2}", cell_text):
                cell_cursor.ParaAdjust = CENTER
        if alignment in alignments:
            cell_cursor.ParaAdjust = alignments[alignment]

        if options.get("first_column_bold", False) and is_first_column:
            cell_cursor.CharWeight = BOLD

        border_color = options.get("border_color")
        if border_color is not None:
            try:
                border = uno.createUnoStruct("com.sun.star.table.BorderLine2")
                border.Color = self.parse_color(border_color)
                border.LineWidth = int(options.get("border_width", 1) * 35.2778)
                for border_name in ("TopBorder", "BottomBorder", "LeftBorder", "RightBorder"):
                    setattr(cell, border_name, border)
            except Exception as e:
                raise RuntimeError("Unable to set table border") from e

    def format_table(self, table, options):
        """Format all cells and optionally the first row as a header."""
        self._insert_table_caption(table, options)
        if "split" in options:
            table.Split = options["split"]
        if "keep_together" in options:
            table.KeepTogether = options["keep_together"]
        for merge in options.get("merge_cells", []):
            try:
                table_cursor = table.createCursorByCellName(merge["start"])
                table_cursor.gotoCellByName(merge["end"], True)
                table_cursor.mergeRange()
            except Exception as e:
                raise RuntimeError("Unable to merge table cells") from e

        header = options.get("header", options.get("format_header", False))
        if options.get("repeat_header", False):
            table.RepeatHeadline = True

        column_widths = options.get("column_widths")
        if column_widths:
            try:
                separators = list(table.TableColumnSeparators)
                position = 0
                for index, width in enumerate(column_widths[:-1]):
                    position += float(width)
                    if index < len(separators):
                        separators[index].Position = int(position * 10)
                        separators[index].IsVisible = True
                table.TableColumnSeparators = tuple(separators)
            except Exception as e:
                raise RuntimeError("Unable to set table column widths") from e

        row_height = options.get("row_height")
        if row_height is not None:
            try:
                height = int(float(row_height) * 35.2778)
                for index in range(table.Rows.getCount()):
                    row = table.Rows.getByIndex(index)
                    row.IsAutoHeight = False
                    row.Height = height
            except Exception as e:
                raise RuntimeError("Unable to set table row height") from e

        for cell_name in table.getCellNames():
            row_number = self._cell_row(cell_name)
            cell_options = dict(options)
            if options.get("zebra", False) and row_number and row_number % 2 == 0:
                cell_options["cell_background"] = options.get("zebra_background", "F2F2F2")
            self._format_table_cell(
                table.getCellByName(cell_name),
                cell_options,
                bool(header and row_number == 1),
                bool(re.match(r"^A\d+$", cell_name)),
            )
            
        
    
  
def precheck_format_request(format_request, fmt):
    """Check target existence and ambiguous searches before changing the document."""
    def check_paragraph_styles(value):
        if not isinstance(value, dict):
            return
        style_name = value.get("paragraph_style")
        if style_name and not fmt.available_paragraph_style(style_name):
            raise ValueError(f"Paragraph style does not exist: {style_name}")
        for nested in value.values():
            check_paragraph_styles(nested)

    check_paragraph_styles(format_request)
    for scope, options in format_request.items():
        if scope == "selection" and fmt.get_selection_cursor() is None:
            raise ValueError("Select text before applying a selection instruction")
        if re.fullmatch(r"paragraph_\d+", str(scope)):
            if fmt.get_paragraph_cursor(int(scope.split("_")[1])) is None:
                raise ValueError(f"Target {scope} does not exist")
        if scope == "paragraph_text":
            fmt.find_paragraph_cursor(options["text"])
        if scope == "paragraphs":
            for paragraph_scope, paragraph_style in options.items():
                if paragraph_scope == "contains":
                    fmt.find_paragraph_cursor(paragraph_style["text"])
                elif fmt.get_paragraph_cursor(int(paragraph_scope.split("_")[1])) is None:
                    raise ValueError(f"Target {paragraph_scope} does not exist")
        table_scopes = options if scope == "tables" else {scope: options} if re.fullmatch(r"table_\d+", str(scope)) else {}
        for table_scope, table_options in table_scopes.items():
            tables = fmt.get_tables(table_scope)
            if not tables:
                raise ValueError(f"Target {table_scope} does not exist")
            for table in tables:
                names = set(table.getCellNames())
                for cell_scope in table_options.get("cells", {}):
                    row, col = map(int, re.fullmatch(r"row_(\d+)_col_(\d+)", cell_scope).groups())
                    if f"{fmt._column_name(col)}{row}" not in names:
                        raise ValueError(f"Target {table_scope}.{cell_scope} does not exist")
                for merge in table_options.get("merge_cells", []):
                    if merge["start"] not in names or merge["end"] not in names:
                        raise ValueError(f"Merge target in {table_scope} does not exist")


def execute_format_request(format_request, fmt):
    report = {"processed": [], "skipped": []}
    if not format_request:
        return report

    completed = report["processed"]
    for page_key, page_value in format_request.items():
        if page_key == "structure":
            count = fmt.format_document_structure(page_value)
            if count:
                completed.append(f"structure ({count} paragraphs)")
            else:
                report["skipped"].append("structure (no matching paragraphs)")
            continue

        if re.fullmatch(r"paragraph_\d+", str(page_key)):
            paragraph_cursor = fmt.get_paragraph_cursor(int(page_key.split("_")[1]))
            if paragraph_cursor is None:
                raise ValueError(f"Target {page_key} no longer exists")
            apply_styles(fmt, paragraph_cursor, page_value)
            completed.append(page_key)
            continue

        if page_key == "paragraph_text":
            paragraph_cursor = fmt.find_paragraph_cursor(page_value.get("text", ""))
            if paragraph_cursor is None:
                raise ValueError("Paragraph text target no longer exists")
            apply_styles(fmt, paragraph_cursor, page_value.get("format", {}))
            completed.append("paragraph_text")
            continue

        if page_key == "paragraphs":
            for paragraph_key, paragraph_style in page_value.items():
                if re.fullmatch(r"paragraph_\d+", str(paragraph_key)):
                    paragraph_cursor = fmt.get_paragraph_cursor(int(paragraph_key.split("_")[1]))
                    if paragraph_cursor is None:
                        raise ValueError(f"Target {paragraph_key} no longer exists")
                    apply_styles(fmt, paragraph_cursor, paragraph_style)
                    completed.append(paragraph_key)
                elif paragraph_key == "contains" and isinstance(paragraph_style, dict):
                    paragraph_cursor = fmt.find_paragraph_cursor(paragraph_style.get("text", ""))
                    if paragraph_cursor is None:
                        raise ValueError("Paragraph text target no longer exists")
                    apply_styles(fmt, paragraph_cursor, paragraph_style.get("format", {}))
                    completed.append("paragraphs.contains")
            continue

        if page_key == "tables":
            for table_key, table_options in page_value.items():
                for table in fmt.get_tables(table_key):
                    fmt.format_table(table, table_options)
                    completed.append(table_key)
                    for cell_key, cell_style in table_options.get("cells", {}).items():
                        cell_match = re.fullmatch(r"row_(\d+)_col_(\d+)", str(cell_key))
                        if cell_match:
                            cell_cursor = fmt.get_table_cell_cursor(
                                table, int(cell_match.group(1)), int(cell_match.group(2))
                            )
                            if cell_cursor is None:
                                raise ValueError(f"Target {table_key}.{cell_key} no longer exists")
                            apply_styles(fmt, cell_cursor, cell_style)
                            completed.append(f"{table_key}.{cell_key}")
            continue

        if str(page_key).startswith("table_"):
            for table in fmt.get_tables(page_key):
                fmt.format_table(table, page_value)
                completed.append(page_key)
            continue

        if page_key == "selection":
            cursor = fmt.get_selection_cursor()
            if cursor is None:
                raise ValueError("Selection no longer exists")
            apply_styles(fmt, cursor, page_value)
            completed.append("selection")
            continue

        if page_key in ["all_pages", "document", "entire_doc"]:
            cursor = fmt.get_document_cursor()
            apply_styles(fmt, cursor, page_value)
            completed.append(page_key)
            continue
            
        raise ValueError(f"Unsupported formatting scope: {page_key}")
    return report


def apply_styles(fmt_instance, target_cursor, line_style_dict):
    """
    在指定的 cursor 上应用具体的样式属性
    :param fmt_instance: Format 类的实例 (用于调用 set_bold 等方法)
    :param target_cursor: 当前操作的 LibreOffice TextCursor 对象
    :param line_style_dict: 具体的样式字典, 如 {"bold": true, "font_color": "FF0000"}
    """
    # 建立指令与类方法的映射映射
    FORMAT_FUNCTION_MAP = {
        "bold": "set_bold",
        "italic": "set_italic",
        "underline": "set_underline",
        "font_size": "set_font_size",
        "font_color": "set_font_color",
        "font_name": "set_font_name",  
        "font_family": "set_font_name", 
        "highlight": "highlight",
        "remove_highlight": "remove_highlight",
        "align_center": "align_center",
        "align_left": "align_left",
        "align_right": "align_right",
        "align_justify": "align_justify",
        "replace_text": "replace_selection",
        "insert_text": "insert_text_at_cursor", # 插入文本
        "clear_format": "clear_format",
        "title": "format_title",
        "body": "format_body",
        "paragraph_style": "set_paragraph_style",
        "first_line_indent": "set_first_line_indent"
    }

    # 特殊处理：替换文本
    if "replace_text" in line_style_dict:
        new_text = line_style_dict["replace_text"]
        fmt_instance.replace_selection(target_cursor, new_text)

    # 特殊处理：文本插入 (因为它有额外的参数 insert_before)
    if "insert_text" in line_style_dict:
        text_to_insert = line_style_dict["insert_text"]
        is_before = line_style_dict.get("insert_before", False)
        fmt_instance.insert_text_at_cursor(target_cursor, text_to_insert, insert_before=is_before)

    # 遍历字典执行其他操作
    for operation, value in line_style_dict.items():
        # 跳过已经处理过的插入指令或逻辑控制键
        if operation in ["insert_text", "insert_before", "replace_text"]:
            continue
            
        if operation in FORMAT_FUNCTION_MAP:
            func_name = FORMAT_FUNCTION_MAP[operation]
            # 从 fmt 实例中获取对应的方法
            func = getattr(fmt_instance, func_name)

            try:
                # 定义不需要参数的方法名
                no_param_actions = [
                    "clear_format", "remove_highlight",
                    "align_center", "align_left", "align_right", "align_justify",
                    "title", "body"
                ]
                
                if operation in no_param_actions:
                    # 如果大模型返回 "bold": true，则执行
                    if value is not False: 
                        func(target_cursor)
                else:
                    # 需要传参的方法 (如颜色、字号)
                    func(target_cursor, value)
                    
            except Exception as e:
                raise RuntimeError(f"Failed to apply {operation}") from e


def execute_undoable_format(format_request, fmt, undo_manager):
    precheck_format_request(format_request, fmt)
    previous_actions = tuple(undo_manager.getAllUndoActionTitles())
    undo_manager.enterUndoContext("writer.ai formatting")
    try:
        report = execute_format_request(format_request, fmt)
    except Exception:
        undo_manager.leaveUndoContext()
        current_actions = tuple(undo_manager.getAllUndoActionTitles())
        if current_actions != previous_actions and current_actions[0] == "writer.ai formatting":
            undo_manager.undo()
        raise
    else:
        undo_manager.leaveUndoContext()
        return report
                
                
class FormatRequestCallback(unohelper.Base, XCallback):
    """Move the completed background request back onto LibreOffice's UI queue."""

    def __init__(self, job, target_doc, generation):
        self.job = job
        self.target_doc = target_doc
        self.generation = generation

    def notify(self, data):
        self.job._finish_format_request(self.target_doc, data, self.generation)


class ConnectionTestCallback(unohelper.Base, XCallback):
    def __init__(self, status_control):
        self.status_control = status_control

    def notify(self, data):
        try:
            self.status_control.getModel().Label = data
        except Exception:
            pass


class ConnectionTestListener(unohelper.Base, XActionListener):
    def __init__(self, job, controls):
        self.job = job
        self.controls = controls
        self._callback = None
        self._async_callback = None

    def disposing(self, event):
        pass

    def actionPerformed(self, event):
        settings = self.job._read_dialog_config(self.controls)
        status = self.controls["connection_status"]
        status.getModel().Label = "Testing connection..."
        self._callback = ConnectionTestCallback(status)
        self._async_callback = self.job.sm.createInstanceWithContext(
            "com.sun.star.awt.AsyncCallback", self.job.ctx
        )
        callback = self._callback
        async_callback = self._async_callback

        def worker():
            errors = []
            try:
                result = self.job.call_model(
                    'Return JSON: {"all_pages":{"bold":true}}',
                    settings.get("api_key", ""), settings.get("model", ""),
                    settings.get("endpoint", ""), error_report=errors,
                )
            except Exception as error:
                log_to_console(f"Connection test failed: {error}")
                result = None
            message = "Connection successful" if result else (
                errors[0] if errors else "Connection test failed"
            )
            try:
                async_callback.addCallback(callback, message)
            except Exception as error:
                log_to_console(f"Unable to display connection test result: {error}")

        threading.Thread(target=worker, name="writer-ai-connection-test", daemon=True).start()


class BackendItemListener(unohelper.Base, XItemListener):
    def __init__(self, backend_control, model_control, endpoint_label, endpoint_control, presets):
        self.backend_control = backend_control
        self.model_control = model_control
        self.endpoint_label = endpoint_label
        self.endpoint_control = endpoint_control
        self.presets = presets

    def itemStateChanged(self, event):
        try:
            selected = self.backend_control.getModel().SelectedItems
            if not selected:
                return
            index = int(selected[0])
            if not 0 <= index < len(self.presets):
                return
            _, model, endpoint = self.presets[index]
            is_custom = index == len(self.presets) - 1
            self.endpoint_label.setVisible(is_custom)
            self.endpoint_control.setVisible(is_custom)
            if model:
                self.model_control.getModel().Text = model
            if endpoint:
                self.endpoint_control.getModel().Text = endpoint
        except Exception as e:
            log_to_console(f"Unable to apply provider preset: {e}")

    def disposing(self, event):
        pass


class MainJob(unohelper.Base, XJobExecutor):
    PASSWORD_URL = "vnd.writer.ai/api"
    PASSWORD_USER = "api_key"
    REQUEST_TIMEOUT = 30

    def __init__(self, ctx):
        log_to_console("MainJob.__init__ called.")

        self.ctx = ctx
        self._request_in_progress = False
        self._request_callback = None
        self._async_callback = None
        self._request_generation = 0
        self._status_indicator = None
        self._api_key_cache = ""
        self._request_timeout_timer = None
        self._request_guard = None

        try:
            self.sm = ctx.getServiceManager()

            self.desktop = self.sm.createInstanceWithContext(
                "com.sun.star.frame.Desktop",
                ctx
            )

        except Exception as e:
            log_to_console(f"Failed to initialize Desktop: {e}")
            raise

    def _finish_format_request(self, target_doc, result_json, generation):
        if generation != self._request_generation:
            log_to_console("Ignoring cancelled or obsolete formatting response.")
            return
        self._request_generation += 1
        self._request_in_progress = False
        self._request_callback = None
        self._async_callback = None
        if self._request_timeout_timer is not None:
            self._request_timeout_timer.cancel()
            self._request_timeout_timer = None
        guard = self._request_guard
        self._request_guard = None
        try:
            response = json.loads(result_json)
            if isinstance(response, dict) and "request" in response:
                format_request = response.get("request", {})
                validation_warnings = response.get("warnings", [])
                request_errors = response.get("errors", [])
            else:
                format_request = response
                validation_warnings = []
                request_errors = []
            if validation_warnings:
                format_request = {}
                request_errors = [
                    "The model returned unsupported formatting instructions; no changes were applied."
                ] + validation_warnings[:5]
            if not format_request:
                message = "Unable to complete the formatting request."
                if request_errors:
                    message += "\n\n" + "\n".join(request_errors)
                elif validation_warnings:
                    message += "\n\nThe model returned an unsupported formatting plan. Please try again."
                else:
                    message += "\n\nThe model returned no executable formatting instructions."
                self._message_box(target_doc, "Writer.AI Request Error", message).execute()
                log_to_console("Formatting was not completed because the model returned no instructions.")
                return

            if guard is None or not guard.can_apply(format_request):
                log_to_console("Formatting plan discarded because the document or selection changed.")
                try:
                    self._message_box(target_doc, "Writer.AI Document Changed",
                                      "The document or selection changed while the request was running. Please submit the instruction again.").execute()
                except Exception:
                    pass
                return

            fmt = Format(self.ctx, target_doc)
            precheck_format_request(format_request, fmt)
            if not self._preview_format_plan(target_doc, format_request, fmt):
                log_to_console("Formatting plan declined by user.")
                return
            if not guard.can_apply(format_request):
                log_to_console("Document changed while the plan was being reviewed.")
                return
            undo_manager = target_doc.getUndoManager()
            report = execute_undoable_format(format_request, fmt, undo_manager)
            log_to_console(
                f"Formatting completed successfully: {len(report['processed'])} target groups."
            )
            try:
                self._show_result(target_doc, report, count_format_operations(format_request))
            except Exception as display_error:
                log_to_console(f"Unable to display formatting result: {display_error}")
        except Exception as e:
            log_to_console(f"Formatting failed after API response: {e}")
            traceback.print_exc(file=sys.stderr)
            try:
                self._message_box(target_doc, "Writer.AI Formatting Error",
                                  f"Formatting failed: {str(e)[:300]}\n\nThe changes were rolled back where possible.").execute()
            except Exception:
                pass
        finally:
            if guard is not None:
                guard.close()
            self._end_status_indicator()
            _finish_queued_format_request()

    def _start_status_indicator(self, target_doc):
        self._end_status_indicator()
        try:
            frame = target_doc.getCurrentController().getFrame()
            indicator = frame.createStatusIndicator()
            indicator.start("Writer.AI: analysing", 100)
            self._status_indicator = indicator
        except Exception as e:
            # Headless documents and older frames may not expose a status bar.
            self._status_indicator = None
            log_to_console(f"Unable to show request status indicator: {e}")

    def _end_status_indicator(self):
        indicator = getattr(self, "_status_indicator", None)
        self._status_indicator = None
        if indicator is not None:
            try:
                indicator.end()
            except Exception as e:
                log_to_console(f"Unable to clear request status indicator: {e}")

    def _password_container(self):
        return self.sm.createInstanceWithContext(
            "com.sun.star.task.PasswordContainer", self.ctx
        )

    def _password_interaction_handler(self):
        return self.sm.createInstanceWithContext(
            "com.sun.star.task.InteractionHandler", self.ctx
        )

    def get_api_key(self):
        if getattr(self, "_api_key_cache", ""):
            return self._api_key_cache
        container = self._password_container()
        handlers = [None]
        try:
            handlers.append(self._password_interaction_handler())
        except Exception as e:
            log_to_console(f"Unable to create password interaction handler: {e}")
        for handler in handlers:
            try:
                record = container.findForName(
                    self.PASSWORD_URL, self.PASSWORD_USER, handler
                )
                if record.UserList and record.UserList[0].Passwords:
                    self._api_key_cache = record.UserList[0].Passwords[0]
                    return self._api_key_cache
            except Exception as e:
                log_to_console(f"Unable to read API key from LibreOffice password storage: {e}")
        return ""

    def set_api_key(self, api_key):
        container = self._password_container()
        self._api_key_cache = api_key or ""
        try:
            container.removePersistent(self.PASSWORD_URL, self.PASSWORD_USER)
        except Exception:
            pass
        if not api_key:
            return
        try:
            handler = self._password_interaction_handler()
            container.allowPersistentStoring(True)
            try:
                container.useDefaultMasterPassword(handler)
            except Exception:
                # LibreOffice may already have an active master-password setup.
                pass
            container.addPersistent(
                self.PASSWORD_URL, self.PASSWORD_USER, (api_key,), handler
            )
        except Exception as e:
            log_to_console(f"Unable to persist API key in LibreOffice password storage: {e}")
            try:
                container.add(self.PASSWORD_URL, self.PASSWORD_USER, (api_key,), None)
                log_to_console("API key saved for the current LibreOffice session.")
            except Exception as session_error:
                self._api_key_cache = ""
                log_to_console(f"Unable to save API key for the current session: {session_error}")

    def _message_box(self, target_doc, title, message):
        frame = target_doc.getCurrentController().getFrame()
        parent_window = frame.getContainerWindow() if frame else None
        toolkit = self.sm.createInstanceWithContext("com.sun.star.awt.Toolkit", self.ctx)
        return toolkit.createMessageBox(parent_window, "errorbox", BUTTONS_OK, title, message)

    def _show_result(self, target_doc, report, planned_properties):
        frame = target_doc.getCurrentController().getFrame()
        parent_window = frame.getContainerWindow() if frame else None
        toolkit = self.sm.createInstanceWithContext("com.sun.star.awt.Toolkit", self.ctx)
        completed = report["processed"]
        skipped = report["skipped"]
        details = "\n".join(completed[:20])
        if len(completed) > 20:
            details += f"\n… and {len(completed) - 20} more processed target groups"
        if skipped:
            details += "\n\nNo matching target:\n" + "\n".join(skipped[:10])
        box = toolkit.createMessageBox(parent_window, "infobox", BUTTONS_OK,
                                       "Writer.AI Completed",
                                       f"Processed target groups: {len(completed)}.\n"
                                       f"No matching target groups: {len(skipped)}.\n"
                                       f"Planned properties: {planned_properties}.\n\n{details}")
        box.execute()

    def _preview_format_plan(self, target_doc, plan, fmt):
        """Show the complete validated plan before touching the document."""
        dialog = self.sm.createInstanceWithContext("com.sun.star.awt.UnoControlDialog", self.ctx)
        model = self.sm.createInstanceWithContext("com.sun.star.awt.UnoControlDialogModel", self.ctx)
        dialog.setModel(model)
        dialog.setTitle("Writer.AI Formatting Plan")
        dialog.setPosSize(0, 0, 600, 540, SIZE)

        def add(name, control_type, x, y, width, height, **properties):
            control_model = model.createInstance("com.sun.star.awt.UnoControl" + control_type + "Model")
            model.insertByName(name, control_model)
            for key, value in properties.items():
                setattr(control_model, key, value)
            dialog.getControl(name).setPosSize(x, y, width, height, POSSIZE)

        warning = "This plan may change text or merge cells. Review it carefully.\n" if changes_document_content(plan) else ""
        add("intro", "FixedText", 10, 10, 580, 38,
            Label=warning + f"Review {count_format_operations(plan)} planned properties before applying.")
        add("targets", "Edit", 10, 52, 580, 64,
            Text="\n".join(describe_plan_targets(plan, fmt)),
            MultiLine=True, VScroll=True, ReadOnly=True)
        try:
            differences = describe_plan_differences(plan, fmt)
        except Exception as error:
            log_to_console(f"Unable to calculate paragraph preview: {error}")
            differences = ["Paragraph differences unavailable; review the full plan below."]
        add("differences", "Edit", 10, 122, 580, 154,
            Text="\n".join(differences), MultiLine=True, VScroll=True,
            HScroll=True, ReadOnly=True)
        add("plan", "Edit", 10, 282, 580, 206,
            Text=json.dumps(plan, ensure_ascii=False, indent=2),
            MultiLine=True, VScroll=True, HScroll=True, ReadOnly=True)
        add("apply", "Button", 200, 499, 90, 26,
            Label="Apply", PushButtonType=OK, DefaultButton=True)
        add("cancel", "Button", 310, 499, 90, 26,
            Label="Cancel", PushButtonType=CANCEL)

        try:
            frame = target_doc.getCurrentController().getFrame()
            parent_window = frame.getContainerWindow() if frame else None
            toolkit = self.sm.createInstanceWithContext("com.sun.star.awt.Toolkit", self.ctx)
            dialog.createPeer(toolkit, parent_window)
            return bool(dialog.execute())
        finally:
            dialog.dispose()

    def _start_format_request(self, target_doc, query, api_key, model, endpoint):
        guard = DocumentGuard(target_doc)
        with _FORMAT_REQUEST_QUEUE_LOCK:
            _FORMAT_REQUEST_QUEUE.append(
                (self, target_doc, query, api_key, model, endpoint, guard)
            )
            queued = _FORMAT_REQUEST_ACTIVE or len(_FORMAT_REQUEST_QUEUE) > 1
        if queued:
            log_to_console("Formatting request queued; waiting for the previous request.")
        _run_next_format_request()

    def _launch_format_request(self, target_doc, query, api_key, model, endpoint, guard):
        if self._request_in_progress:
            log_to_console("A formatting request is already running.")
            guard.close()
            _finish_queued_format_request()
            return

        if not guard.can_apply({}):
            log_to_console("Queued formatting request skipped because the document changed.")
            try:
                self._message_box(target_doc, "Writer.AI Document Changed",
                                  "The document changed before analysis began. Please submit the instruction again.").execute()
            except Exception:
                pass
            guard.close()
            _finish_queued_format_request()
            return

        async_callback = self.sm.createInstanceWithContext(
            "com.sun.star.awt.AsyncCallback", self.ctx
        )
        self._request_generation += 1
        generation = self._request_generation
        callback = FormatRequestCallback(self, target_doc, generation)
        self._request_callback = callback
        self._async_callback = async_callback
        self._request_in_progress = True
        self._request_guard = guard
        self._start_status_indicator(target_doc)

        def request_timeout():
            try:
                async_callback.addCallback(callback, json.dumps({
                    "request": {},
                    "warnings": [],
                    "errors": ["The model request timed out after 30 seconds."],
                    "timeout": True,
                }))
            except Exception as e:
                log_to_console(f"Unable to deliver request timeout to Writer: {e}")

        self._request_timeout_timer = threading.Timer(
            self.REQUEST_TIMEOUT + 1, request_timeout
        )
        self._request_timeout_timer.daemon = True
        self._request_timeout_timer.start()

        def request_worker():
            validation_report = []
            request_errors = []
            try:
                result = self.call_model(
                    query, api_key, model, endpoint, validation_report, request_errors
                )
                result_json = json.dumps({
                    "request": result or {},
                    "warnings": validation_report,
                    "errors": request_errors,
                })
            except Exception as e:
                log_to_console(f"Background model request failed: {e}")
                request_errors.append("The background request failed. Please try again later.")
                result_json = json.dumps({
                    "request": {},
                    "warnings": validation_report,
                    "errors": request_errors,
                })
            try:
                async_callback.addCallback(callback, result_json)
            except Exception as e:
                log_to_console(f"Unable to deliver model response to Writer: {e}")

        threading.Thread(target=request_worker, name="writer-ai-api", daemon=True).start()
        log_to_console("Formatting request started in background; analysing document.")

    def get_config(self, key, default):
        name_file = "writerai.json"
        path_settings = self.sm.createInstanceWithContext('com.sun.star.util.PathSettings', self.ctx)
        user_config_path = getattr(path_settings, "UserConfig")
        if user_config_path.startswith('file://'):
            user_config_path = str(uno.fileUrlToSystemPath(user_config_path))
        config_file_path = os.path.join(user_config_path, name_file)
        if not os.path.exists(config_file_path):
            return default
        try:
            with open(config_file_path, 'r') as file:
                config_data = json.load(file)
        except (IOError, json.JSONDecodeError):
            return default
        return config_data.get(key, default)

    def set_config(self, key, value):
        name_file = "writerai.json"
        path_settings = self.sm.createInstanceWithContext('com.sun.star.util.PathSettings', self.ctx)
        user_config_path = getattr(path_settings, "UserConfig")
        if user_config_path.startswith('file://'):
            user_config_path = str(uno.fileUrlToSystemPath(user_config_path))
        config_file_path = os.path.join(user_config_path, name_file)
        config_data = {}
        if os.path.exists(config_file_path):
            try:
                with open(config_file_path, 'r') as file:
                    config_data = json.load(file)
            except (IOError, json.JSONDecodeError):
                pass
        config_data[key] = value
        try:
            with open(config_file_path, 'w') as file:
                json.dump(config_data, file, indent=4)
        except IOError as e:
            log_to_console(f"Error writing to config: {e}")

    def _as_bool(self, value):
        if isinstance(value, str):
            return value.lower() in ('true', '1', 't', 'y', 'yes')
        return bool(value)

    BACKEND_PRESETS = [
        ("Alibaba Cloud Bailian - Kimi K3", "kimi-k3", "https://dashscope.aliyuncs.com/compatible-mode/v1"),
        ("Moonshot - Kimi K3", "kimi-k3", "https://api.moonshot.ai/v1"),
        ("Custom OpenAI-compatible API", "", ""),
    ]
    BUILTIN_TEMPLATES = (
        (
            "Academic Paper",
            "Format this document as a formal academic paper. Apply a clear title, "
            "consistent headings, professional body text, and first-line indentation "
            "for body paragraphs.",
        ),
        (
            "Official Document",
            "Format this document as a formal official document. Use a clear title, "
            "consistent numbered headings, readable body text, and consistent paragraph spacing.",
        ),
        (
            "Resume",
            "Format this document as a professional resume. Create a clear title, "
            "consistent section headings, compact body text, and clean alignment.",
        ),
    )
    
    @staticmethod
    def call_model(query, api_key, model, endpoint, validation_report=None, error_report=None):
        """
    使用 OpenAI 兼容的 Chat Completions 接口生成格式化指令。
    
    Args:
        user_input: 用户输入的自然语言，如 "将第一页第一行进行黄色高亮标记"
        api_key: 服务商 API Key
        model: 模型名称
        endpoint: Base URL 或完整 Chat Completions URL
        
    Returns:
        dict: 结构化后的指令字典。若解析失败则返回空字典。
    """
        system_prompt = """
You are Writer.AI, a LibreOffice Writer formatting assistant. Return ONLY one
JSON object in the exact format below. Do not return a generic {"scope": ...,
"properties": ...} object. For example, "bold the first paragraph" is
{"paragraph_1":{"bold":true}}, and "highlight the first paragraph" is
{"paragraph_1":{"highlight":true}}. A paragraph is not a visual line.
Use paragraph_N for a numbered paragraph. Visual page and line positions are
not supported because they change with layout. If the user provides only a
visual page and line, return {} rather than guessing a paragraph. Valid scopes: selection,
all_pages, paragraph_N, paragraph_text, paragraphs, structure, tables, and
table_N. Valid properties include
bold, italic, underline, font_size, font_color, font_name, highlight,
remove_highlight, align_center, align_left, align_right, align_justify, title,
body, paragraph_style, first_line_indent, replace_text, insert_text,
insert_before, and table properties header, header_background, cell_background,
table_background, header_bold, align, repeat_header, first_column_bold, zebra,
zebra_background, row_height, border_color, border_width, column_widths,
merge_cells, auto_align, caption, number, caption_prefix, split, and
keep_together. Use paragraph_N or paragraph_text for stable paragraph targets.
Use paragraph_text only for a distinctive phrase; repeated matches are rejected
by the executor, so prefer paragraph_N when the user gives an ordinal.
Use tables.table_N.cells.row_N_col_N for table cells. Use structure with title,
heading_1, heading_2, heading_3, and body for document roles. Values must match
their property meaning. Preserve the document's text, punctuation, formulas,
code, and citations. Only use replace_text, insert_text, caption, or merge_cells
when the user explicitly asks to change content or table structure. Never
include explanations or Markdown.
"""
        if not api_key:
            api_key = os.environ.get("WRITER_AI_API_KEY", "")
        parsed_endpoint = urlparse(endpoint)
        host = (parsed_endpoint.hostname or "").lower()
        local_endpoint = host in {"localhost", "127.0.0.1", "::1"}
        if (not model or not endpoint or not host or
                parsed_endpoint.scheme not in {"http", "https"} or
                (parsed_endpoint.scheme != "https" and not local_endpoint) or
                (not api_key and not local_endpoint)):
            log_to_console("A valid HTTPS endpoint, model, and API key are required for remote providers.")
            if error_report is not None:
                error_report.append("Check the model, endpoint, and API key in Settings. Remote endpoints must use HTTPS; local endpoints may omit the key.")
            return {}

        endpoint = endpoint.rstrip("/")
        if not endpoint.endswith("/chat/completions"):
            endpoint += "/chat/completions"

        request_body = {
            "model": model,
            "messages": [
                {"role": "system", "content": system_prompt},
                {"role": "user", "content": query},
            ],
            "stream": False,
            "max_completion_tokens": 512,
        }
        if model.lower().startswith("qwen") and (
            host == "dashscope.aliyuncs.com" or host == "dashscope-intl.aliyuncs.com"
            or host.endswith(".maas.aliyuncs.com")
        ):
            request_body["response_format"] = {"type": "json_object"}
        payload = json.dumps(request_body).encode("utf-8")
        headers = {"Content-Type": "application/json"}
        if api_key and not local_endpoint:
            headers["Authorization"] = f"Bearer {api_key}"
        request = Request(
            endpoint,
            data=payload,
            headers=headers,
            method="POST",
        )

        try:
            with urlopen(request, timeout=MainJob.REQUEST_TIMEOUT) as response:
                response_data = json.loads(response.read().decode("utf-8"))
            content = response_data["choices"][0]["message"]["content"]
            try:
                data = validate_format_request(parse_model_json(content), validation_report)
                log_to_console(f"Structured query: {data}")
                return data
            except (ValueError, TypeError) as e:
                log_to_console(f"JSON Parsing Error: {e}")
                if error_report is not None:
                    error_report.append("The model returned an unreadable formatting plan. Please try again.")
                return None
        except HTTPError as e:
            error_body = e.read().decode("utf-8", errors="replace")
            log_to_console(f"Model API 请求失败: HTTP {e.code}: {error_body}")
            if error_report is not None:
                error_report.append(f"The model service returned HTTP {e.code}. Check the API key, model, and endpoint.")
            return {}
        except (URLError, KeyError, IndexError, json.JSONDecodeError) as e:
            log_to_console(f"Model API 调用失败: {e}")
            if error_report is not None:
                error_report.append("Could not connect to the model service or received an invalid response. Check the network and settings.")
            return {}


    
    def _detect_backend(self):
        # ... [Unchanged] ...
        model_name = self.get_config("model", "").lower()
        endpoint = self.get_config("endpoint", "").lower().rstrip("/")
        for i, preset in enumerate(self.BACKEND_PRESETS):
            if (
                preset[1]
                and preset[1].lower() == model_name
                and (not endpoint or preset[2].lower().rstrip("/") == endpoint)
            ):
                return i
        return len(self.BACKEND_PRESETS) - 1

    def _read_dialog_config(self, controls):
        # ... [Unchanged] ...
        result = {}
        initial = getattr(self, "_settings_initial_values", {})
        if "backend" in controls and controls["backend"].getModel().SelectedItems:
            backend_idx = controls["backend"].getModel().SelectedItems[0]
            preset = self.BACKEND_PRESETS[backend_idx]
            result["model"] = controls["model"].getModel().Text or preset[1]
            result["endpoint"] = controls["endpoint"].getModel().Text or preset[2]
        if "api_key" in controls:
            result["api_key"] = controls["api_key"].getModel().Text or initial.get("api_key", "")
        if "model" in controls:
            result["model"] = controls["model"].getModel().Text or initial.get("model", "")
        if "endpoint" in controls:
            result["endpoint"] = controls["endpoint"].getModel().Text or initial.get("endpoint", "")
        return result

    def _save_settings(self, result):
        if not result:
            log_to_console("No settings to save.")
            return
        api_key = result.pop("api_key", None)
        if api_key is not None:
            self.set_api_key(api_key)
        for key, value in result.items():
            self.set_config(key, value)
        log_to_console("Settings saved.")

    def _custom_templates(self):
        templates = self.get_config("templates", [])
        if not isinstance(templates, list):
            return []
        return [
            (item["name"], item["instruction"])
            for item in templates
            if isinstance(item, dict)
            and isinstance(item.get("name"), str)
            and item["name"].strip()
            and isinstance(item.get("instruction"), str)
            and item["instruction"].strip()
        ]

    def _all_templates(self):
        return list(self.BUILTIN_TEMPLATES) + self._custom_templates()

    def _save_template(self, name, instruction):
        name = str(name).strip()
        instruction = str(instruction).strip()
        if not name or not instruction:
            return
        templates = [
            {"name": item_name, "instruction": item_instruction}
            for item_name, item_instruction in self._custom_templates()
            if item_name.lower() != name.lower()
        ]
        templates.append({"name": name, "instruction": instruction})
        self.set_config("templates", templates)
        log_to_console(f"Template saved: {name}")

    def templates_box(self, title="Templates"):
        templates = self._all_templates()
        if not templates:
            return None
        dialog = self.sm.createInstanceWithContext("com.sun.star.awt.UnoControlDialog", self.ctx)
        model = self.sm.createInstanceWithContext("com.sun.star.awt.UnoControlDialogModel", self.ctx)
        dialog.setModel(model)
        dialog.setTitle(title)
        dialog.setPosSize(0, 0, 520, 220, SIZE)

        list_model = model.createInstance("com.sun.star.awt.UnoControlListBoxModel")
        model.insertByName("templates", list_model)
        list_model.StringItemList = tuple(item[0] for item in templates)
        list_model.SelectedItems = (0,)
        list_model.MultiSelection = False
        list_control = dialog.getControl("templates")
        list_control.setPosSize(10, 10, 500, 150, POSSIZE)

        apply_model = model.createInstance("com.sun.star.awt.UnoControlButtonModel")
        model.insertByName("apply", apply_model)
        apply_model.Label = "Apply Template"
        apply_model.PushButtonType = OK
        apply_control = dialog.getControl("apply")
        apply_control.setPosSize(130, 175, 125, 26, POSSIZE)

        cancel_model = model.createInstance("com.sun.star.awt.UnoControlButtonModel")
        model.insertByName("cancel", cancel_model)
        cancel_model.Label = "Cancel"
        cancel_model.PushButtonType = CANCEL
        cancel_control = dialog.getControl("cancel")
        cancel_control.setPosSize(265, 175, 125, 26, POSSIZE)

        frame = self.desktop.getCurrentFrame()
        window = frame.getContainerWindow() if frame else None
        dialog.createPeer(self.sm.createInstanceWithContext("com.sun.star.awt.Toolkit", self.ctx), window)
        try:
            if not dialog.execute():
                return None
            selected = list_model.SelectedItems
            if not selected or not 0 <= selected[0] < len(templates):
                return None
            return templates[selected[0]][1]
        finally:
            dialog.dispose()


    def settings_box(self, title="", x=None, y=None):
        log_to_console("--- Starting settings_box ---")
        WIDTH, HEIGHT = 700, 260
        HORI_MARGIN, VERT_MARGIN = 10, 10
        BUTTON_WIDTH, BUTTON_HEIGHT = 170, 26
        HORI_SEP, VERT_SEP = 10, 10
        LABEL_WIDTH, LABEL_HEIGHT, EDIT_HEIGHT = 150, 20, 24
        
        ctx = self.ctx
        def create(name):
            log_to_console(f"  Creating service: {name}")
            return ctx.getServiceManager().createInstanceWithContext(name, ctx)

        try:
            dialog = create("com.sun.star.awt.UnoControlDialog")
            dialog_model = create("com.sun.star.awt.UnoControlDialogModel")
            log_to_console("Dialog and model created.")
            
            dialog.setModel(dialog_model)
            dialog.setTitle(title)
            dialog.setPosSize(0, 0, WIDTH, HEIGHT, SIZE)
            log_to_console("Dialog model set, title and size set.")

            def add(name, ctrl_type, x, y, width, height, props):
                log_to_console(f"  Adding control '{name}' of type '{ctrl_type}'")
                model = dialog_model.createInstance("com.sun.star.awt.UnoControl" + ctrl_type + "Model")
                dialog_model.insertByName(name, model)
                control = dialog.getControl(name)
                control.setPosSize(x, y, width, height, POSSIZE)
                for key, value in props.items():
                    setattr(model, key, value)
                return control

            controls = {}
            edit_width = WIDTH - HORI_MARGIN * 2 - LABEL_WIDTH - HORI_SEP
            
            y_pos = VERT_MARGIN
            add("label_backend", "FixedText", HORI_MARGIN, y_pos + 4, LABEL_WIDTH, LABEL_HEIGHT, {"Label": "Model Preset:"})
            backend_names = tuple(p[0] for p in self.BACKEND_PRESETS)
            current_backend_idx = self._detect_backend()
            controls["backend"] = add("list_backend", "ListBox", HORI_MARGIN + LABEL_WIDTH, y_pos,
                edit_width, EDIT_HEIGHT,
                {"Dropdown": True, "StringItemList": backend_names, "SelectedItems": (current_backend_idx,)})

            y_pos += EDIT_HEIGHT + VERT_SEP
            add("label_api_key", "FixedText", HORI_MARGIN, y_pos + 4, LABEL_WIDTH, LABEL_HEIGHT, {"Label": "API Key:"})
            controls["api_key"] = add("edit_api_key", "Edit", HORI_MARGIN + LABEL_WIDTH, y_pos,
                edit_width, EDIT_HEIGHT, {"Text": self.get_api_key(), "EchoChar": ord("*")})

            y_pos += EDIT_HEIGHT + VERT_SEP
            preset = self.BACKEND_PRESETS[current_backend_idx]
            self._settings_initial_values = {
                "api_key": self.get_api_key(),
                "model": str(self.get_config("model", preset[1]) or preset[1]),
                "endpoint": str(self.get_config("endpoint", preset[2]) or preset[2]),
            }
            add("label_model", "FixedText", HORI_MARGIN, y_pos + 4, LABEL_WIDTH, LABEL_HEIGHT, {"Label": "Model:"})
            controls["model"] = add("edit_model", "Edit", HORI_MARGIN + LABEL_WIDTH, y_pos,
                edit_width, EDIT_HEIGHT, {"Text": str(self.get_config("model", preset[1]))})

            y_pos += EDIT_HEIGHT + VERT_SEP
            controls["endpoint_label"] = add("label_endpoint", "FixedText", HORI_MARGIN, y_pos + 4, LABEL_WIDTH, LABEL_HEIGHT, {"Label": "Base URL (advanced):"})
            controls["endpoint"] = add("edit_endpoint", "Edit", HORI_MARGIN + LABEL_WIDTH, y_pos,
                edit_width, EDIT_HEIGHT, {"Text": str(self.get_config("endpoint", preset[2]))})
            is_custom_backend = current_backend_idx == len(self.BACKEND_PRESETS) - 1
            controls["endpoint_label"].setVisible(is_custom_backend)
            controls["endpoint"].setVisible(is_custom_backend)

            controls["backend_listener"] = BackendItemListener(
                controls["backend"], controls["model"], controls["endpoint_label"],
                controls["endpoint"], self.BACKEND_PRESETS
            )
            controls["backend"].addItemListener(controls["backend_listener"])

            controls["connection_status"] = add(
                "connection_status", "FixedText", HORI_MARGIN, HEIGHT - 78,
                WIDTH - HORI_MARGIN * 2, LABEL_HEIGHT, {"Label": ""}
            )
            y_pos = HEIGHT - BUTTON_HEIGHT - VERT_MARGIN
            button_start_x = (WIDTH - (BUTTON_WIDTH * 3 + HORI_SEP * 2)) / 2
            test_button = add("btn_test", "Button", button_start_x, y_pos,
                              BUTTON_WIDTH, BUTTON_HEIGHT, {"Label": "Test Connection"})
            controls["test_listener"] = ConnectionTestListener(self, controls)
            test_button.addActionListener(controls["test_listener"])
            add("btn_ok", "Button", button_start_x + BUTTON_WIDTH + HORI_SEP, y_pos, BUTTON_WIDTH, BUTTON_HEIGHT, {
                "PushButtonType": OK,
                "DefaultButton": True,
                "Label": "Save Settings",
                "HelpText": "Save the API key, model, and provider settings.",
            })
            add("btn_cancel", "Button", button_start_x + BUTTON_WIDTH * 2 + HORI_SEP * 2, y_pos, BUTTON_WIDTH, BUTTON_HEIGHT, {
                "PushButtonType": CANCEL,
                "Label": "Cancel",
                "HelpText": "Close without changing saved settings.",
            })

            log_to_console("All controls added.")

            frame = self.desktop.getCurrentFrame()
            window = frame.getContainerWindow() if frame else None
            if not window:
                log_to_console("ERROR: Could not get window to create dialog.")
                return {}
            
            log_to_console("About to create peer.")
            dialog.createPeer(create("com.sun.star.awt.Toolkit"), window)
            log_to_console("Peer created.")
            
            ret = {}
            log_to_console("About to execute dialog.")
            if dialog.execute():
                log_to_console("Dialog executed, OK pressed.")
                ret = self._read_dialog_config(controls)
            else:
                log_to_console("Dialog executed, Cancel pressed.")
            
        except Exception as e:
            log_to_console("--- EXCEPTION in settings_box ---")
            log_to_console(e)
            traceback.print_exc(file=sys.stderr)
            ret = {}
        finally:
            log_to_console("Finally block: Disposing dialog.")
            if 'dialog' in locals() and dialog:
                dialog.dispose()
        
        log_to_console("--- Exiting settings_box ---")
        return ret

    def input_box(self, message, title="", default="", x=None, y=None):
        # ... [Unchanged] ...
        WIDTH = 500
        HORI_MARGIN = 10
        VERT_MARGIN = 10
        BUTTON_WIDTH = 80
        BUTTON_HEIGHT = 25
        HORI_SEP = 10
        VERT_SEP = 10
        LABEL_HEIGHT = 25
        EDIT_HEIGHT = 25 * 5

        HEIGHT = VERT_MARGIN * 3 + LABEL_HEIGHT + VERT_SEP + EDIT_HEIGHT + BUTTON_HEIGHT

        ctx = self.ctx
        def create(name):
            return ctx.getServiceManager().createInstanceWithContext(name, ctx)
        
        dialog = create("com.sun.star.awt.UnoControlDialog")
        dialog_model = create("com.sun.star.awt.UnoControlDialogModel")
        dialog.setModel(dialog_model)
        dialog.setVisible(False)
        dialog.setTitle(title)
        dialog.setPosSize(0, 0, WIDTH, HEIGHT, SIZE)

        def add(name, ctrl_type, x, y, width, height, props):
            model = dialog_model.createInstance("com.sun.star.awt.UnoControl" + ctrl_type + "Model")
            dialog_model.insertByName(name, model)
            control = dialog.getControl(name)
            control.setPosSize(x, y, width, height, POSSIZE)
            for key, value in props.items():
                setattr(model, key, value)

        add("label", "FixedText", HORI_MARGIN, VERT_MARGIN, WIDTH - HORI_MARGIN * 2, LABEL_HEIGHT, {"Label": str(message)})
        
        edit_y = VERT_MARGIN + LABEL_HEIGHT + VERT_SEP
        edit_width = WIDTH - HORI_MARGIN * 2
        add("edit", "Edit", HORI_MARGIN, edit_y, edit_width, EDIT_HEIGHT, 
            {"Text": str(default), "MultiLine": True, "VScroll": True})
        
        buttons_y = edit_y + EDIT_HEIGHT + VERT_SEP

        ok_x = (WIDTH - (BUTTON_WIDTH * 2 + HORI_SEP)) / 2
        add("btn_ok", "Button", ok_x, buttons_y, BUTTON_WIDTH, BUTTON_HEIGHT, {"PushButtonType": OK, "DefaultButton": True})
        cancel_x = ok_x + BUTTON_WIDTH + HORI_SEP
        add("btn_cancel", "Button", cancel_x, buttons_y, BUTTON_WIDTH, BUTTON_HEIGHT, {"PushButtonType": CANCEL})

        frame = self.desktop.getCurrentFrame()
        window = frame.getContainerWindow() if frame else None
        dialog.createPeer(create("com.sun.star.awt.Toolkit"), window)
        
        edit = dialog.getControl("edit")
        edit.setFocus()
        
        ret = edit.getModel().Text if dialog.execute() else ""
        dialog.dispose()
        return ret


    def trigger(self, args):
        log_to_console(f"\n--- Trigger called with args: {args} ---")

        if args == "setting":
            log_to_console("Entering settings branch...")
            try:
                result = self.settings_box("Writer.ai Settings")
                self._save_settings(result)
            except Exception as e:
                log_to_console("--- EXCEPTION in trigger(setting) ---")
                log_to_console(e)
                traceback.print_exc(file=sys.stderr)

        elif args == "format":
            log_to_console("Entering format branch...")
            try:
                # 1. Initialize UNO environment
                smgr = self.ctx.getServiceManager() # 建议使用 self.ctx
                desktop = smgr.createInstanceWithContext("com.sun.star.frame.Desktop", self.ctx)
                
                # 2. Ask user for the source document
                active_frame = desktop.getCurrentFrame()
                parent_window = active_frame.getContainerWindow() if active_frame else None
                
                msg_box = smgr.createInstanceWithContext("com.sun.star.awt.Toolkit", self.ctx).createMessageBox(
                    parent_window, 
                    "querybox", 
                    BUTTONS_YES_NO, 
                    "AI Formatter", 
                    "Would you like to format the CURRENTLY active document?\n\n(Select 'No' to pick a different file.)"
                )
                
                choice = msg_box.execute()
                target_doc = None

                if choice == YES:
                    target_doc = desktop.getCurrentComponent()
                    log_to_console("Mode: Processing active document.")
                else:
                    file_url = pick_writer_file(self.ctx)
                    if file_url:
                        # Open the selected file
                        target_doc = desktop.loadComponentFromURL(file_url, "_blank", 0, ())
                        log_to_console(f"Mode: Opened new document {file_url}")
                    else:
                        log_to_console("User cancelled file selection.")
                        return # Exit gracefully

                # 3. Validation: Ensure it's a Writer document
                if not target_doc or not hasattr(target_doc, "supportsService") or \
                   not target_doc.supportsService("com.sun.star.text.TextDocument"):
                    log_to_console("Error: Selected component is not a Writer document.")
                    return

                # 4. Get User Input for AI
                user_input = self.input_box(
                    "Document Format (e.g. Bold all headings):",
                    "AI Formatter",
                )

                if not user_input:
                    log_to_console("User cancelled input.")
                    return

                # 5. Start the API request without blocking LibreOffice's UI thread.
                api_key = self.get_api_key() or os.environ.get("WRITER_AI_API_KEY", "")
                model = self.get_config("model", "kimi-k3")
                endpoint = self.get_config("endpoint", "https://dashscope.aliyuncs.com/compatible-mode/v1")
                self._start_format_request(target_doc, user_input, api_key, model, endpoint)

            except Exception as e:
                log_to_console("--- EXCEPTION in trigger(format) ---")
                log_to_console(str(e))
                traceback.print_exc(file=sys.stderr)
                
g_ImplementationHelper = unohelper.ImplementationHelper()
g_ImplementationHelper.addImplementation(
    MainJob,
    "org.extension.writerai.do",
    ("com.sun.star.task.Job",),
)
log_to_console("Script loaded, implementation added.")
