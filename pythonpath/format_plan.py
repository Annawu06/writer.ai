"""Parse and validate model formatting plans without LibreOffice dependencies."""

import json
import math
import re


FORMAT_KEYS = {
    "bold", "italic", "underline", "font_size", "font_color",
    "font_name", "font_family", "highlight", "remove_highlight",
    "align_center", "align_left", "align_right", "align_justify",
    "replace_text", "insert_text", "insert_before", "clear_format",
    "title", "body", "paragraph_style", "first_line_indent",
}
TABLE_KEYS = {
    "cell_background", "table_background", "header", "format_header",
    "header_background", "header_bold", "align", "font_name",
    "font_family", "font_size", "font_color", "repeat_header",
    "first_column_bold", "zebra", "zebra_background", "row_height",
    "border_color", "border_width", "column_widths",
    "merge_cells", "auto_align",
    "caption", "title", "number", "caption_prefix", "split", "keep_together",
}


def _valid_value(key, value):
    if key in {"bold", "italic", "remove_highlight", "clear_format", "title", "body", "insert_before", "header", "format_header", "header_bold", "repeat_header", "first_column_bold", "zebra", "auto_align", "split", "keep_together"}:
        return isinstance(value, bool)
    if key in {"font_size", "first_line_indent", "row_height", "border_width", "number"}:
        limits = {
            "font_size": (1, 200), "first_line_indent": (0, 20),
            "row_height": (0, 1000), "border_width": (0, 20),
            "number": (0, 100000),
        }
        low, high = limits[key]
        return (isinstance(value, (int, float)) and not isinstance(value, bool)
                and math.isfinite(value) and low <= value <= high)
    if key == "column_widths":
        return (
            isinstance(value, list)
            and len(value) >= 2
            and all(isinstance(item, (int, float)) and not isinstance(item, bool)
                    and math.isfinite(item) and 0 < item <= 100 for item in value)
            and abs(sum(value) - 100) < 0.01
        )
    if key == "merge_cells":
        return (
            isinstance(value, list)
            and all(
                isinstance(item, dict)
                and isinstance(item.get("start"), str)
                and isinstance(item.get("end"), str)
                and re.fullmatch(r"[A-Z]{1,3}[1-9]\d{0,4}", item.get("start", ""))
                and re.fullmatch(r"[A-Z]{1,3}[1-9]\d{0,4}", item.get("end", ""))
                for item in value
            )
        )
    if key == "underline":
        return isinstance(value, bool) or (isinstance(value, str) and
            bool(re.fullmatch(r"(?:[0-9]|1[0-8])(?:[0-9A-Fa-f]{6})?", value)))
    if key in {"font_color", "highlight", "cell_background", "table_background", "header_background", "zebra_background", "border_color"}:
        return (key == "highlight" and value is True) or (
            isinstance(value, int) and not isinstance(value, bool) and 0 <= value <= 0xFFFFFF
        ) or (isinstance(value, str) and bool(re.fullmatch(r"#?[0-9A-Fa-f]{6}", value)))
    if key == "align":
        return value in {"left", "right", "center", "justify"} if isinstance(value, str) else False
    if key in {"font_name", "font_family", "replace_text", "insert_text", "paragraph_style", "caption", "caption_prefix"}:
        return isinstance(value, str) and len(value) <= 10000
    if key.startswith("align_"):
        return isinstance(value, bool)
    return False


def _validation_issue(report, message):
    if report is not None:
        report.append(message)


def parse_model_json(content):
    """Read one JSON object, including a fenced or prose-wrapped response.

    Keep repairs narrow: a guessed comma or quote can change a formatting
    instruction, so malformed plans other than trailing commas are rejected.
    """
    if isinstance(content, dict):
        return content
    if not isinstance(content, str):
        raise ValueError("The model response is not text or an object")
    decoder = json.JSONDecoder()
    for match in re.finditer(r"\{", content):
        candidate = content[match.start():]
        for attempt in (candidate, _remove_trailing_commas(candidate)):
            try:
                result, _ = decoder.raw_decode(attempt)
            except json.JSONDecodeError:
                continue
            if isinstance(result, dict):
                return result
    raise ValueError("The model response contains no valid JSON object")


def _remove_trailing_commas(text):
    chars = []
    in_string = False
    escaped = False
    for index, char in enumerate(text):
        if in_string:
            chars.append(char)
            if escaped:
                escaped = False
            elif char == "\\":
                escaped = True
            elif char == '"':
                in_string = False
        elif char == '"':
            in_string = True
            chars.append(char)
        elif char == "," and re.match(r"\s*[}\]]", text[index + 1:]):
            continue
        else:
            chars.append(char)
    return "".join(chars)


def validate_style(style, allowed_keys, report=None):
    if not isinstance(style, dict):
        _validation_issue(report, "Ignoring formatting properties because the value is not an object")
        return {}
    clean_style = {}
    for key, value in style.items():
        if key not in allowed_keys:
            _validation_issue(report, f"Ignoring unsupported formatting property: {key}")
            continue
        if (isinstance(value, str) if key == "title" and allowed_keys is TABLE_KEYS
                else _valid_value(key, value)):
            clean_style[key] = value
        else:
            _validation_issue(report, f"Ignoring invalid value for formatting property: {key}")
    if allowed_keys is FORMAT_KEYS:
        conflicting = False
        alignment = [key for key in ("align_left", "align_right", "align_center", "align_justify")
                     if clean_style.get(key) is True]
        if len(alignment) > 1:
            _validation_issue(report, "Conflicting paragraph alignments in one target")
            conflicting = True
        paragraph_styles = [key for key in ("title", "body", "paragraph_style")
                            if clean_style.get(key)]
        if len(paragraph_styles) > 1:
            _validation_issue(report, "Conflicting paragraph styles in one target")
            conflicting = True
        if conflicting:
            return {}
    return clean_style


def validate_format_request(request, report=None):
    """Keep only known Writer scopes and properties before document mutation."""
    if not isinstance(request, dict):
        _validation_issue(report, "Formatting response must be a JSON object")
        return {}

    clean_request = {}
    for scope, value in request.items():
        if scope in {"selection", "all_pages", "document", "entire_doc"}:
            clean_style = validate_style(value, FORMAT_KEYS, report)
            if clean_style:
                clean_request[scope] = clean_style
        elif re.fullmatch(r"paragraph_\d+", str(scope)):
            clean_style = validate_style(value, FORMAT_KEYS, report)
            if clean_style:
                clean_request[scope] = clean_style
        elif scope == "paragraph_text":
            if isinstance(value, dict) and isinstance(value.get("text"), str):
                clean_style = validate_style(value.get("format", {}), FORMAT_KEYS, report)
                if clean_style:
                    clean_request[scope] = {"text": value["text"], "format": clean_style}
        elif scope == "paragraphs":
            if not isinstance(value, dict):
                continue
            paragraphs = {}
            for paragraph_scope, paragraph_style in value.items():
                if re.fullmatch(r"paragraph_\d+", str(paragraph_scope)):
                    clean_style = validate_style(paragraph_style, FORMAT_KEYS, report)
                    if clean_style:
                        paragraphs[paragraph_scope] = clean_style
                elif paragraph_scope == "contains" and isinstance(paragraph_style, dict) and isinstance(paragraph_style.get("text"), str):
                    clean_style = validate_style(paragraph_style.get("format", {}), FORMAT_KEYS, report)
                    if clean_style:
                        paragraphs[paragraph_scope] = {"text": paragraph_style["text"], "format": clean_style}
            if paragraphs:
                clean_request[scope] = paragraphs
        elif scope == "structure":
            if not isinstance(value, dict):
                continue
            structure = {}
            if isinstance(value.get("infer"), bool):
                structure["infer"] = value["infer"]
            for role in {"title", "heading_1", "heading_2", "heading_3", "body"}:
                if role in value:
                    role_style = value[role]
                    if role_style is True:
                        structure[role] = True
                    elif isinstance(role_style, dict):
                        clean_style = validate_style(role_style, FORMAT_KEYS, report)
                        if clean_style:
                            structure[role] = clean_style
            if len(structure) > 1 or any(role in structure for role in {"title", "heading_1", "heading_2", "heading_3", "body"}):
                clean_request[scope] = structure
        elif scope == "tables":
            if not isinstance(value, dict):
                continue
            tables = {}
            for table_scope, table_style in value.items():
                if re.fullmatch(r"(?:table_\d+|table_all|all)", str(table_scope)):
                    table_base = {
                        key: item for key, item in table_style.items()
                        if key != "cells"
                    } if isinstance(table_style, dict) else table_style
                    clean_style = validate_style(table_base, TABLE_KEYS, report)
                    if isinstance(table_style, dict) and isinstance(table_style.get("cells"), dict):
                        cells = {}
                        for cell_scope, cell_style in table_style["cells"].items():
                            if re.fullmatch(r"row_\d+_col_\d+", str(cell_scope)):
                                cell_clean_style = validate_style(cell_style, FORMAT_KEYS, report)
                                if cell_clean_style:
                                    cells[cell_scope] = cell_clean_style
                        if cells:
                            clean_style["cells"] = cells
                    if clean_style:
                        tables[table_scope] = clean_style
            if tables:
                clean_request[scope] = tables
        elif re.fullmatch(r"table_\d+", str(scope)):
            clean_style = validate_style(value, TABLE_KEYS, report)
            if clean_style:
                clean_request[scope] = clean_style
        else:
            _validation_issue(report, f"Ignoring unsupported formatting scope: {scope}")
    if report is not None and clean_request != request and not report:
        _validation_issue(report, "Formatting plan contains unsupported or invalid structure")
    return clean_request


def count_format_operations(request):
    """Count requested style properties and implicit structure style changes."""
    if not isinstance(request, dict):
        return 0

    def table_properties(style):
        return (len(style) - ("cells" in style)
                + sum(len(cell_style) for cell_style in style.get("cells", {}).values()))

    count = 0
    for scope, value in request.items():
        if scope == "paragraph_text":
            count += len(value["format"])
        elif scope == "paragraphs":
            count += sum(len(style["format"]) if target == "contains" else len(style)
                         for target, style in value.items())
        elif scope == "structure":
            count += sum(1 + (len(style) if isinstance(style, dict) else 0)
                         for role, style in value.items() if role != "infer")
        elif scope == "tables":
            count += sum(table_properties(style) for style in value.values())
        elif re.fullmatch(r"table_\d+", scope):
            count += table_properties(value)
        else:
            count += len(value)
    return count


def changes_document_content(plan):
    if isinstance(plan, dict):
        return any(key in {"replace_text", "insert_text", "merge_cells", "caption"}
                   or (key == "title" and isinstance(value, str))
                   or changes_document_content(value) for key, value in plan.items())
    if isinstance(plan, list):
        return any(changes_document_content(value) for value in plan)
    return False
