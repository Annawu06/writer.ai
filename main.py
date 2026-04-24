import uno
import re
import unohelper
import dashscope # 需安装：pip install dashscope
from http import HTTPStatus
import json
import os
import traceback
import sys # To print to stderr
from com.sun.star.task import XJobExecutor
from com.sun.star.awt import MessageBoxButtons as MSG_BUTTONS
from com.sun.star.awt import XActionListener, XItemListener
from com.sun.star.awt.PosSize import POS, SIZE, POSSIZE
from com.sun.star.awt.PushButtonType import OK, CANCEL
from com.sun.star.util.MeasureUnit import TWIP

from com.sun.star.awt.FontWeight import BOLD
from com.sun.star.awt.FontSlant import ITALIC
from com.sun.star.awt import FontUnderline
from com.sun.star.style.ParagraphAdjust import LEFT, RIGHT, CENTER, BLOCK

from com.sun.star.ui.dialogs.TemplateDescription import FILEOPEN_SIMPLE

from com.sun.star.awt.MessageBoxButtons import BUTTONS_YES_NO
from com.sun.star.awt.MessageBoxResults import YES

from com.sun.star.awt.FontWeight import NORMAL
from com.sun.star.awt.FontSlant import NONE
from com.sun.star.awt.FontUnderline import NONE as UNDERLINE_NONE


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


def get_doc(ctx):
    smgr = ctx.getServiceManager()
    desktop = smgr.createInstanceWithContext(
        "com.sun.star.frame.Desktop", ctx
    )
    return desktop.getCurrentComponent()
    
       
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
        
    def get_all_lines_cursor(self, page_num):

        try:
            # 1. go to page 
            self.goto_page(page_num)
            view_cursor = self.doc.CurrentController.getViewCursor()
            
            # 2. move to the start of the page
            view_cursor.jumpToStartOfPage()
            start_range = view_cursor.getStart()
            
            # 3. move to the end of the page
            view_cursor.jumpToEndOfPage()
            end_range = view_cursor.getEnd()
            
            # 4. create a TextCursor
            cursor = self.doc.Text.createTextCursorByRange(start_range)
            cursor.gotoRange(end_range, True)
            return cursor
        except Exception as e:
            log_to_console(f"Error creating page cursor: {e}")
            return self.doc.Text.createTextCursor()

    def get_selection(self):
        """
                获取选中的文本对象
        """
        selection = self.controller.getSelection()
        if selection.getCount() > 0:
            return selection.getByIndex(0)
        return None
        
    def goto_page(self, page):
        view_cursor = self.get_cursor()
        view_cursor.jumpToPage(page)
        view_cursor.jumpToStartOfPage() 

    def goto_line(self, line):
        view_cursor = self.get_cursor()
        cursor = self.doc.Text.createTextCursorByRange(view_cursor.getStart())
        for _ in range(line - 1):
            if not cursor.gotoNextParagraph(False):
                break
                
        cursor.gotoStartOfParagraph(False)
        cursor.gotoEndOfParagraph(True)
        return cursor
        
        
    def find_paragraphs_by_styles(doc, target_styles=None):
        """
        输入:
            doc: LibreOffice 文档对象 (XComponent)
            target_styles: 想要匹配的样式名称列表 (list)
        输出:
            matches: 包含匹配段落对象的列表
        """
        if target_styles is None:
            target_styles = ["Title", "Subtitle", "Text body"] 
            target_styles.extend([f"Heading {i}" for i in range(1, 11)])

        matches = []
        
        # 遍历文档中的所有内容
        paragraphs = doc.Text.createEnumeration()
        
        while paragraphs.hasMoreElements():
            para = paragraphs.nextElement()

            # 确保当前元素是一个段落（过滤掉表格等其他对象）
            if para.supportsService("com.sun.star.text.Paragraph"):
                # 检查段落样式是否在目标列表中
                if para.ParaStyleName in target_styles:
                    matches.append({
                        "style": para.ParaStyleName,
                        "text": para.String,
                        "object": para  
                    })
                    

        return matches



    # ------------------------------------------------
    # Text Style
    # ------------------------------------------------

    def set_bold(self,cursor, value=True):
        cursor = cursor
        cursor.CharWeight = BOLD

    def set_italic(self,cursor, value=True):
        cursor = cursor
        cursor.CharPosture = ITALIC

    def set_underline(self, cursor, value):
        try:
            print(f"value:{value}")

            val_str = str(value).strip()
            print(f"val_str{val_str}")
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

            print(f"Fixed Debug - Raw: {val_str}, Style: {cursor.CharUnderline}, HasColor: {cursor.CharUnderlineHasColor}")

        except Exception as e:
            print(f"Critical Underline Error: {e}")
            cursor.CharUnderline = 0


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
                log_to_console(f"Error setting font name: {e}")



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
            log_to_console(f"Error setting color: {e}")


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

        cursor.gotoStartOfParagraph(False)
        cursor.gotoEndOfParagraph(True)

        cursor.ParaAdjust = CENTER

    def align_left(self, cursor):
        cursor.gotoStartOfParagraph(False)
        cursor.gotoEndOfParagraph(True)
        cursor.ParaAdjust = LEFT

    def align_right(self, cursor):
        cursor.gotoStartOfParagraph(False)
        cursor.gotoEndOfParagraph(True)
        cursor.ParaAdjust = RIGHT

    def align_justify(self, cursor):
        cursor.gotoStartOfParagraph(False)
        cursor.gotoEndOfParagraph(True)
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
                    return self.doc.Text.createTextCursorByRange(selected_range)
            except Exception as e:
                log_to_console(f"Error getting selection cursor: {e}")
            
            # 修复点：确保这里也使用 self.doc 或调用已有的获取全篇游标的方法
            return self.get_document_cursor()

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
            log_to_console(f"Error in insert_text_at_cursor: {e}")


    def replace_selection(self, cursor, text):
        try:
            if cursor:
                # 这里的逻辑不需要 doc，直接操作 cursor 即可
                cursor.setString(text)
        except Exception as e:
            log_to_console(f"Error in replace_selection: {e}")

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
            
        
    
  
def execute_format_request(format_request, fmt):
    if not format_request:
        return

    for page_key, page_value in format_request.items():
        # --- 核心修改点 ---
        # 如果 key 是 selection，或者 page_value 包含特定指示
        if page_key == "selection":
            cursor = fmt.get_selection_cursor() # 你需要在 Format 类实现这个方法
            apply_styles(fmt, cursor, page_value)
            continue

        if page_key in ["all_pages", "document", "entire_doc"]:
            # 如果 AI 脑抽在 all_pages 里提到了 selection，优先处理选区
            if "selection" in str(page_value).lower():
                cursor = fmt.get_selection_cursor()
            else:
                cursor = fmt.get_document_cursor()
            
            apply_styles(fmt, cursor, page_value)
            continue
            
        # 2. 按页处理逻辑 (page_n)
        try:
            if "_" not in page_key:
                continue
                
            page_num = int(page_key.split("_")[1])
            fmt.goto_page(page_num)

            for line_key, line_value in page_value.items():
                # 确定每一行的 Cursor 范围
                if line_key in ["line_all", "all"]:
                    cursor = fmt.get_all_lines_cursor(page_num)
                else:
                    try:
                        line_num = int(line_key.split("_")[1])
                        cursor = fmt.goto_line(line_num)
                    except (ValueError, IndexError):
                        continue
                
                # 调用具体的样式应用函数
                apply_styles(fmt, cursor, line_style_dict=line_value)

        except Exception as e:
            log_to_console(f"Error processing page {page_key}: {e}")
            
            
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
        "clear_format": "clear_format"
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
                    "bold", "italic", "clear_format", "remove_highlight",
                    "align_center", "align_left", "align_right", "align_justify"
                ]
                
                if operation in no_param_actions:
                    # 如果大模型返回 "bold": true，则执行
                    if value is not False: 
                        func(target_cursor)
                else:
                    # 需要传参的方法 (如颜色、字号)
                    func(target_cursor, value)
                    
            except Exception as e:
                log_to_console(f"Error executing {operation} on cursor: {e}")
                
                
class MainJob(unohelper.Base, XJobExecutor):
    def __init__(self, ctx):
        log_to_console("MainJob.__init__ called.")

        self.ctx = ctx

        try:
            self.sm = ctx.getServiceManager()

            self.desktop = self.sm.createInstanceWithContext(
                "com.sun.star.frame.Desktop",
                ctx
            )

        except Exception as e:
            log_to_console(f"Failed to initialize Desktop: {e}")
            raise

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
        ("Gemini 3 Pro", "chat", "https://generativelanguage.googleapis.com/v1beta"),
        ("Gemini 3 Flash", "chat", "https://generativelanguage.googleapis.com/v1beta"),
        ("QWen", "chat", "https://dashscope.aliyuncs.com/api/v1/services/aigc/text-generation/generation"),
    ]
    
    @staticmethod
    def askQwen(query,api_key = "sk-f361ac282d2044d1a9523413ee925382"):
        """
    使用 Qwen 大模型将自然语言指令转换为 LibreOffice Writer 的结构化配置字典。
    
    Args:
        user_input: 用户输入的自然语言，如 "将第一页第一行进行黄色高亮标记"
        api_key: 阿里云 DashScope 的 API Key
        
    Returns:
        dict: 结构化后的指令字典。若解析失败则返回空字典。
    """
        print(f"original query is :{query}")
        
        # 1. 构造系统提示词，严格定义输出规范
        system_prompt = ("""
                            # Role
                            You are a formatting expert specifically designed for LibreOffice Writer. Your mission is to translate natural language instructions from users into precise JSON formatting commands.

                            # Output Format
                            You MUST output ONLY the JSON format. Do not include any explanations. The structure is as follows:
                            {
                              "all_pages": { "Property": "Value" },  // For document-wide operations
                              "page_n": {
                                "line_all": { "Property": "Value" }, // For page-wide operations
                                "line_n": { "Property": "Value" }    // For specific line operations
                              }
                            }

                            # 1. Structural Range Protocol (MUTUALLY EXCLUSIVE)
                            You MUST choose ONLY ONE of the following three structures. NEVER mix them.

                            - [RULE A] GLOBAL ONLY: 
                              If the user says "all", "entire doc", "everywhere".
                              Format: {"all_pages": {"bold": true, ...}} 
                              (STRICT: No "line_n" keys allowed inside)

                            - [RULE B] PAGE-WIDE: 
                              If the user says "all of page 1", "entire page 2".
                              Format: {"page_1": {"line_all": {"bold": true, ...}}}

                            - [RULE C] SPECIFIC LOCATION: 
                              If the user says "first paragraph", "line 3", "page 1 line 1".
                              Format: {"page_1": {"line_1": {"bold": true, ...}}}
                              (STRICT: Even if it's the "first paragraph" of the whole doc, use "page_1" -> "line_1", NEVER "all_pages")

                            # 2. Logic Priority
                            - If a specific location (paragraph/line) is mentioned, RULE C overrides everything.
                            - NEVER nest "line_n" or "line_all" under "all_pages".

                            # 3. Property Assignment Logic (Strict Rules)

                            ## A. Color Rules (font_color / highlight / underline)
                            - Supports Hex strings (e.g., "FF5733").
                            - Supports Semantic Colors: 
                              - "warning" -> Red ("FF0000")
                              - "Tiffany Blue" -> "0ABAB5"
                              - "Sakura Pink" -> "FFB7C5"
                              - "Success/Go" -> "008000"
                              - "Sky Blue" -> "87CEEB"
                              - "Gold" -> "FFD700"

                            ## B. Font Names (font_name)
                            - Use semantic tags: serif, sans-serif, monospace, code, formal, modern.
                            - Chinese Fonts: Use "heiti" (Black-style), "songti" (Standard/Print), "kaiti" (Handwriting).
                            - Specific Fonts: If mentioned (e.g., "Arial", "Consolas"), use the name directly.

                            ## C. Underline Composite Construction
                            - Format: The value of underline MUST follow the [StyleCode][HexColor] format.
                            - CharUnderline style mapping: 
                              {0:NONE, 1:SINGLE, 2:DOUBLE, 3:DOTTED, 5:DASH, 6:LONG_DASH, 7:DASH_DOT, 8:DASH_DOT_DOT, 9:SMALL_WAVE, 10:WAVE, 11:DOUBLE_WAVE, 12:BOLD, 13:BOLD_DOTTED, 14:BOLD_DASH, 15:BOLD_LONG_DASH, 16:BOLD_DASH_DOT, 17:BOLD_DASH_DOT_DOT, 18:BOLD_WAVE}
                            - Assignment Logic:
                              - Style Only: If the user only says "underline" or "wave underline" without a color, output only the Style Code (e.g., "1" or "10"). (Logic: Set CharUnderlineHasColor to False).
                              - Color Only: If the user says "blue underline", output "10000FF" (1 is Single, 0000FF is Blue).
                              - Style & Color: If the user says "red double underline", output "2FF0000" (2 is Double, FF0000 is Red).
                            - Strict Atomic Interpretation: 
                              Interpret the following as single underline types rather than font weight combinations:
                              - bold wave underline = 18
                              - bold dotted underline = 13
                              - bold dash underline = 14
                              - bold dash dot underline = 16
                              *Do not split "bold" into font weight when it appears inside underline style names.*

                            # 4. Property Names (MUST MATCH BACKEND)
                            You must ONLY use these property keys:
                            1. Styles: "bold" (bool), "italic" (bool), "underline" (str/bool), "clear_format" (true).
                            2. Fonts: "font_name" (semantic/literal), "font_size" (number), "font_color" (Hex).
                            3. Highlights: "highlight" (Hex or true for default), "remove_highlight" (true).
                            4. Alignment: "align_center", "align_left", "align_right", "align_justify" (all bool).

                            # 5. Text Insertion & Spatial Localization Protocol
                            If the user wants to add, insert, or format text, follow these STRICT structural rules:

                            - **Structural Hierarchy Priority**: 
                                1. If the user specifies a location (e.g., "second paragraph", "line 5", "page 1"), you MUST use the `page_n` -> `line_n` hierarchy. 
                                2. DO NOT use "all_pages" unless the command explicitly applies to the entire document (e.g., "set whole doc font size to 12").
                                3. Treat "paragraph" as synonymous with "line" (e.g., "second paragraph" maps to "line_2").

                            - **Insertion Parameters**:
                                - "insert_text": (string) The exact text content to be added.
                                - "insert_before": (boolean) 
                                    - `true`: Use for "before", "at the start of", "prefix", "in front of", "at the beginning".
                                    - `false`: Use for "after", "at the end of", "append", "suffix", "behind". (Default)

                            - **Example Mapping**:
                                - User: "insert 'Hello' after the 2nd paragraph"
                                - Target: {"page_1": {"line_2": {"insert_text": "Hello", "insert_before": false}}}
                                
                                - User: "bold the first line"
                                - Target: {"page_1": {"line_1": {"bold": true}}}
                                
                            - **Text Replacement Protocol**:
                            - If the user says "replace X with Y", "change line X to Y", or "overwrite":
                            - Use "replace_text": "Y"
                            - Hierarchy: Ensure it's inside the correct `page_n` -> `line_n`.

                            # 6. Scope Selection Rules (CRITICAL)
                            - **Selection Mode**: 
                                - If the user mentions "selection", "selected part", "what I highlighted", or "this":
                                - YOU MUST use "selection" as the top-level key. 
                                - Example: {"selection": {"replace_text": "new text"}}
                                
                            - **Defaulting Rules**:
                                - If the user says "replace" WITHOUT specifying a line number OR "selection":
                                - DO NOT default to "line_1". 
                                - Instead, ALWAYS use "selection" as the default scope. 
                                - Logic: Users usually want to operate where their cursor is currently blinking.

                            - **Strict All Pages**:
                                - Only use "all_pages" if the user says "whole document", "everything", or "all".

                            # 7. Operational Constraints
                            - NO REDUNDANCY: Do not invent keys like "underline_color".
                            - NO ASSUMPTIONS: Do not add a "highlight" if the user only asked for "underline".
                            - COLOR PRIORITY: If a color is specified for a style (like underline), put that hex code directly into that property instead of adding a separate highlight.

                            # 8. Examples
                            - User: "Change entire doc to dark blue, font size 12"
                              Assistant: {"all_pages": {"font_color": "00008B", "font_size": 12}}

                            - User: "Bold everything on page 1 and use heiti font"
                              Assistant: {"page_1": {"line_all": {"bold": true, "font_name": "heiti"}}}

                            - User: "Set the highlight of page 2 line 4 to Sakura Pink"
                              Assistant: {"page_2": {"line_4": {"highlight": "FFB7C5"}}}

                  """
                
        )
        # 2. 调用 Qwen 模型 (以 qwen-turbo 为例，也可根据需求换成 qwen-max)
        dashscope.api_key = api_key
        #print("1")
        response = dashscope.Generation.call(
            model=dashscope.Generation.Models.qwen_turbo,
            messages=[
                {'role': 'system', 'content': system_prompt},
                {'role': 'user', 'content': query}
            ],
            result_format='message'
        )
        #print("2")

        # 3. 处理响应结果
        if response.status_code == HTTPStatus.OK:
            #print("3")
            content = response.output.choices[0].message.content
            print(f"content:{content}")
            try:
                # 1. 清理字符串
                clean_json = content.replace("```json", "").replace("```", "").strip()
                
                # 2. 只解析一次并存储在变量中
                data = json.loads(clean_json)
                
                # 3. 打印并返回
                log_to_console(f"Structured query: {data}")
                return data
            except Exception as e:
                log_to_console(f"JSON Parsing Error: {e}")
                return None
        else:
            print(f"API 请求失败: {response.code} - {response.message}")
            return {}     


    
    def _detect_backend(self):
        # ... [Unchanged] ...
        model_name = self.get_config("model", "").lower()
        for i, preset in enumerate(self.BACKEND_PRESETS):
            if preset[0].lower() == model_name:
                return i
        return 0

    def _read_dialog_config(self, controls):
        # ... [Unchanged] ...
        result = {}
        if "backend" in controls and controls["backend"].getModel().SelectedItems:
            backend_idx = controls["backend"].getModel().SelectedItems[0]
            preset = self.BACKEND_PRESETS[backend_idx]
            result["model"] = preset[0]
            result["api_type"] = preset[1]
            result["endpoint"] = preset[2]
        if "api_key" in controls:
            result["api_key"] = controls["api_key"].getModel().Text
        return result

    def _save_settings(self, result):
        log_to_console("Saving settings:", result)
        if not result:
            log_to_console("No settings to save.")
            return
        for key, value in result.items():
            self.set_config(key, value)
        log_to_console("Settings saved.")


    def settings_box(self, title="", x=None, y=None):
        log_to_console("--- Starting settings_box ---")
        WIDTH, HEIGHT = 600, 150
        HORI_MARGIN, VERT_MARGIN = 10, 10
        BUTTON_WIDTH, BUTTON_HEIGHT = 100, 26
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
                edit_width, EDIT_HEIGHT, {"Text": str(self.get_config("api_key", ""))})

            y_pos = HEIGHT - BUTTON_HEIGHT - VERT_MARGIN
            button_start_x = (WIDTH - (BUTTON_WIDTH * 2 + HORI_SEP)) / 2
            add("btn_ok", "Button", button_start_x, y_pos, BUTTON_WIDTH, BUTTON_HEIGHT, {"PushButtonType": OK, "DefaultButton": True})
            add("btn_cancel", "Button", button_start_x + BUTTON_WIDTH + HORI_SEP, y_pos, BUTTON_WIDTH, BUTTON_HEIGHT, {"PushButtonType": CANCEL})

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

        from com.sun.star.awt.PosSize import POS, SIZE, POSSIZE
        from com.sun.star.awt.PushButtonType import OK, CANCEL
        
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
        BUTTONS_YES_NO = uno.getConstantByName("com.sun.star.awt.MessageBoxButtons.BUTTONS_YES_NO")
        YES = uno.getConstantByName("com.sun.star.awt.MessageBoxResults.YES")
        
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
                    "Document Format:",
                    "AI Formatter",
                    "Example: Bold all headings or highlight keywords."
                )

                if not user_input:
                    log_to_console("User cancelled input.")
                    return

                # 5. AI Process & Execution
                # Note: Passing target_doc to your formatting logic is CRITICAL
                format_request = MainJob.askQwen(user_input)
                
                # Make sure your Format class is initialized with the CORRECT doc
                fmt = Format(self.ctx, target_doc) 
                
                execute_format_request(format_request, fmt)

                log_to_console("Formatting completed successfully.")

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
