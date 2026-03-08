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
from com.sun.star.awt.FontUnderline import SINGLE
from com.sun.star.style.ParagraphAdjust import LEFT, RIGHT, CENTER, BLOCK

from com.sun.star.ui.dialogs.TemplateDescription import FILEOPEN_SIMPLE

from com.sun.star.awt.MessageBoxButtons import BUTTONS_YES_NO
from com.sun.star.awt.MessageBoxResults import YES

from com.sun.star.awt.FontWeight import NORMAL
from com.sun.star.awt.FontSlant import NONE
from com.sun.star.awt.FontUnderline import NONE as UNDERLINE_NONE



def pick_writer_file(ctx):

    smgr = ctx.getServiceManager()

    file_picker = smgr.createInstanceWithContext(
        "com.sun.star.ui.dialogs.FilePicker",
        ctx
    )

    file_picker.initialize((FILEOPEN_SIMPLE,))

    file_picker.setTitle("Select Writer Document")

    # 文件过滤器
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
    # In the LibreOffice Python context, print often goes to a specific log file.
    # Writing to stderr is sometimes more reliable for seeing output in a console.
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
        # 1. 扩展的标准 Web 颜色映射 (常用部分，你可以根据需要继续增加)
        # 提示：LibreOffice 的颜色是 Long 类型，即 R*65536 + G*256 + B
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

    # 处理 True (AI表示“默认高亮”)
        if color is True:
            return std_colors["yellow"]
        
        if not color or not isinstance(color, str):
            return std_colors["yellow"]

        # 2. 预处理：【关键】除了转小写，还要去掉空格
        # 这样 "Dark Red" 会变成 "darkred"，就能匹配字典了
        clean_color = color.lower().replace(" ", "").strip().lstrip('#')

        # 3. 尝试从字典匹配
        if clean_color in std_colors:
            return std_colors[clean_color]

        # 4. 尝试十六进制匹配 (处理 AI 给出的 "B2C8D9")
        if re.fullmatch(r'[0-9a-f]{3}|[0-9a-f]{6}', clean_color):
            try:
                if len(clean_color) == 3:
                    clean_color = ''.join([c*2 for c in clean_color])
                return int(clean_color, 16)
            except ValueError:
                pass

        # 5. 保底：建议改成紫色 0x800080。
        # 如果运行后看到紫色，说明输入的颜色既不在字典里，也不是合法的十六进制。
        return std_colors["yellow"]
        
    def get_cursor(self):
        """
            获取当前文本光标
        """
        return self.controller.getViewCursor()
        
    def get_all_lines_cursor(self, page_num):
        """
            获取指定页码整页内容的 Cursor
        """
        try:
            # 1. 先跳转到该页
            self.goto_page(page_num)
            view_cursor = self.doc.CurrentController.getViewCursor()
            
            # 2. 移动到该页开头
            view_cursor.jumpToStartOfPage()
            start_range = view_cursor.getStart()
            
            # 3. 移动到该页结尾
            view_cursor.jumpToEndOfPage()
            end_range = view_cursor.getEnd()
            
            # 4. 创建一个包含整个页面的 TextCursor
            cursor = self.doc.Text.createTextCursorByRange(start_range)
            cursor.gotoRange(end_range, True) # True 表示“扩展选中”
            return cursor
        except Exception as e:
            log_to_console(f"Error creating page cursor: {e}")
            return self.doc.Text.createTextCursor() # 出错则返回普通 cursor 兜底

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
        view_cursor.jumpToStartOfPage() # 确保在页首

    def goto_line(self, line):
        view_cursor = self.get_cursor()
        # 创建一个锚定在当前 ViewCursor 位置的逻辑光标
        cursor = self.doc.Text.createTextCursorByRange(view_cursor.getStart())
        
        # 向下移动 (line-1) 次
        for _ in range(line - 1):
            if not cursor.gotoNextParagraph(False):
                break
                
        cursor.gotoStartOfParagraph(False)
        cursor.gotoEndOfParagraph(True)
        return cursor



    # ------------------------------------------------
    # 文本样式
    # ------------------------------------------------

    def set_bold(self,cursor):
        cursor = cursor
        cursor.CharWeight = BOLD

    def set_italic(self,cursor):
        cursor = cursor
        cursor.CharPosture = ITALIC

    def set_underline(self, cursor, color=None):

        cursor.CharUnderline = 1
        cursor.CharUnderlineHasColor = True
        if color:
            cursor.CharUnderlineColor = self.parse_color(color)



    def set_font_name(self, cursor, font_name):
            try:
                if not font_name or not isinstance(font_name, str):
                    return
                
                # 扩展语义映射：将 AI 的描述性词汇映射到 LibreOffice 常用字体
                font_map = {
                    # --- 基础类别 ---
                    "serif": "Libre Serif",
                    "sans-serif": "Libre Sans",
                    "monospace": "Liberation Mono",
                    "code": "Consolas",
                    
                    # --- 现代/简约风格 ---
                    "modern": "Noto Sans",
                    "clean": "DejaVu Sans",
                    "minimal": "Inter",
                    
                    # --- 正式/学术风格 ---
                    "formal": "Libre Baskerville",
                    "academic": "Linux Libertine G",
                    "professional": "Liberation Serif",
                    "classic": "Times New Roman",
                    
                    # --- 中文字体语义 (针对 Debian 环境常用) ---
                    "chinese": "Noto Sans CJK SC",
                    "heiti": "Noto Sans CJK SC",
                    "songti": "Noto Serif CJK SC",
                    "kaiti": "AR PL UKai CN",
                    "microsoft yahei": "Microsoft YaHei",
                    
                    # --- 艺术/手写风格 ---
                    "handwriting": "Comic Sans MS", # 虽然名声不好但很常用
                    "elegant": "Apple Chancery",
                    "title": "Linux Biolinum G"
                }
                
                # 处理逻辑：
                # 1. 尝试全字匹配字典（如 "serif"）
                # 2. 尝试去掉空格后匹配（如 "sansserif"）
                # 3. 如果都不匹配，直接使用原字符串（假设用户输入了具体的字体名如 "Arial"）
                clean_name = font_name.lower().replace(" ", "").replace("-", "")
                target_font = font_map.get(clean_name, font_name)
                
                # 设置三种字符集属性，确保兼容性
                cursor.CharFontName = target_font          # 西文字体
                cursor.CharFontNameAsian = target_font     # 中日韩字体
                cursor.CharFontNameComplex = target_font   # 复杂文字（如阿拉伯语）
                
            except Exception as e:
                log_to_console(f"Error setting font name: {e}")

    # ------------------------------------------------
    # 字体大小
    # ------------------------------------------------

    def set_font_size(self, cursor, size):
        size = float(size)

        cursor.CharHeight = size
        cursor.CharHeightAsian = size

    # ------------------------------------------------
    # 字体颜色
    # ------------------------------------------------

    def set_font_color(self, cursor, rgb):
        try:
            # 如果漏网之鱼是字符串，这里做最后一次转换
            if isinstance(rgb, str):
                rgb = self.parse_color(rgb)
            
            cursor.CharColor = int(rgb) 
        except Exception as e:
            log_to_console(f"Error setting color: {e}")

    # ------------------------------------------------
    # 高亮
    # ------------------------------------------------
    def highlight(self, cursor, color=None):

        if color is None or color is True:
            color = "yellow"

        uno_color = self.parse_color(color)

        cursor.CharBackColor = uno_color
            

    def remove_highlight(self,cursor):
        cursor = cursor
        cursor.CharBackColor = -1

    # ------------------------------------------------
    # 段落对齐
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
    # 插入文本
    # ------------------------------------------------

    def insert_text(self, cursor,text):
        cursor = cursor
        cursor.setString(cursor.getString() + text)

    # ------------------------------------------------
    # 替换选中文本
    # ------------------------------------------------

    def replace_selection(self, cursor,text):
        selection = self.get_selection()
        if selection:
            selection.setString(text)

    # ------------------------------------------------
    # 获取当前文本
    # ------------------------------------------------

    def get_selected_text(self,cursor):
        selection = cursor
        if selection:
            return selection.getString()
        return ""

    # ------------------------------------------------
    # 清除所有格式
    # ------------------------------------------------



    def clear_format(self, cursor):

        cursor.CharWeight = NORMAL
        cursor.CharPosture = NONE
        cursor.CharUnderline = UNDERLINE_NONE

        cursor.CharStrikeout = 0

        cursor.CharColor = -1
        cursor.CharBackColor = -1

        cursor.CharHeight = 12
            
        

    
  
def execute_format_request(format_request, fmt):

    FORMAT_FUNCTION_MAP = {

        # -------------------------
        # 文本样式
        # -------------------------
        "bold": "set_bold",
        "italic": "set_italic",
        "underline": "set_underline",

        # -------------------------
        # 字体
        # -------------------------
        "font_size": "set_font_size",
        "font_color": "set_font_color",
        "font_name": "set_font_name",  # 新增
        "font_family": "set_font_name", # 增加一个别名更稳妥

        # -------------------------
        # 高亮
        # -------------------------
        "highlight": "highlight",
        "remove_highlight": "remove_highlight",

        # -------------------------
        # 段落对齐
        # -------------------------
        "align_center": "align_center",
        "align_left": "align_left",
        "align_right": "align_right",
        "align_justify": "align_justify",

        # -------------------------
        # 文本操作
        # -------------------------
        "insert_text": "insert_text",
        "replace_selection": "replace_selection",

        # -------------------------
        # 文档工具
        # -------------------------
        "clear_format": "clear_format"
    }

    for page_key, page_value in format_request.items():

        page_num = int(page_key.split("_")[1])
                
        fmt.goto_page(page_num)

        for line_key, line_value in page_value.items():
                    # 1. 确定 Cursor 范围
                    if line_key == "line_all" or line_key == "all":
                        # 获取整页 Cursor
                        cursor = fmt.get_all_lines_cursor(page_num)
                    else:
                        try:
                            # 确保 line_1 这种格式能被正确解析
                            line_num = int(line_key.split("_")[1])
                            cursor = fmt.goto_line(line_num)
                        except (ValueError, IndexError):
                            log_to_console(f"Skipping invalid line key: {line_key}")
                            continue

                    # 2. 执行格式化操作 
                    for operation, value in line_value.items():
                        if operation in FORMAT_FUNCTION_MAP:
                            # 动态获取 fmt 类中的方法
                            func_name = FORMAT_FUNCTION_MAP[operation]
                            func = getattr(fmt, func_name)
                            
                            # 3. 颜色预处理：如果是颜色相关操作，先转为 UNO 整数
                            if operation in ["font_color", "highlight", "underline"]:
                                value = fmt.parse_color(value)
                            
                            # 4. 调用函数
                            try:
                                # 兼容处理：有些函数可能不需要参数，有些需要 value
                                if value is True and operation not in ["font_color", "highlight", "underline", "font_name", "font_size"]:
                                    func(cursor) # 处理类似 set_bold 这种开关
                                else:
                                    func(cursor, value)
                            except Exception as e:
                                log_to_console(f"Error executing {operation} on {line_key}: {e}")
                        
  
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
        # ... [Unchanged] ...
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
        # ... [Unchanged] ...
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
                你是一个 LibreOffice Writer 文档格式解析器。

                你的任务是把用户的自然语言编辑指令转换为结构化 JSON。

                输出必须满足以下规则：

                1. 输出必须是一个 **纯 JSON 对象**。
                2. 不允许包含解释文字、Markdown、代码块或注释。
                3. JSON 结构必须为：

                {
                  "page_n": {
                    "line_m": {
                      "property": value
                    }
                  }
                }

                其中：

                page_n = 页码，例如 page_1
                line_m = 行号，例如 line_3

                4. 同一页的多个行必须合并在同一个 page 对象内，不能覆盖！Group instructions by page to avoid duplicate keys，例如：

                正确：

                {
                  "page_1": {
                    "line_1": {"bold": true},
                    "line_4": {"highlight": true}
                  }
                }

                错误（禁止重复 page_1）：

                {
                  "page_1": {"line_1": {"bold": true}},
                  "page_1": {"line_4": {"highlight": true}}
                }

                5. 支持的属性：
                
                font_size
                font_color
                font_name

                align_center
                align_left
                align_right
                align_justify
                bold
                italic
                underline
                highlight
                insert_text
                replace_selection
                clear_format

                6. 属性值规则 (highlight / underline / font_color)：

                - 如果用户没有指定具体颜色：
                                      使用 true（例如 {"highlight": true}），程序将应用默认值。

                                    - 如果用户指定了具体颜色（无论是中文“天蓝色”、英文“cyan”、RGB 还是十六进制）：
                                      你必须将其转换为标准的 6 位十六进制字符串（不带 # 号）。
                                      例如：
                                      输入 "红色" -> 输出 "FF0000"
                                      输入 "light blue" -> 输出 "ADD8E6"
                                      输入 "rgb(255, 0, 0)" -> 输出 "FF0000"
                                      输入 "#00FF00" -> 输出 "00FF00"
                                      如果你识别到品牌特定颜色（如 Tiffany Blue, Coca-Cola Red），请务必使用其公认的标准 Hex 代码（如 Tiffany Blue = 0ABAB5）

                7. 组合逻辑示例：

                                    输入：将第一页第二行设置为紫色高亮并加粗
                                    输出：
                                    {
                                      "page_1": {
                                        "line_2": {
                                          "highlight": "800080",
                                          "bold": true
                                        }
                                      }
                                    }

                8. 严禁事项：
                - 严禁在 JSON 的颜色值中输出 "red"、"blue" 等单词。
                - 必须统一输出 6 位十六进制（如 "FFFF00"），确保后端 parse_color 函数能直接转换。
                10. 如果用户没有指定某个属性，不要推测。

                ---

                                示例

                                输入：

                                bold the first line of page 1 and highlight line 4 on page 1

                                输出：

                                {
                                  "page_1": {
                                    "line_1": {"bold": true},
                                    "line_4": {"highlight": true}
                                  }
                                }

                                输入：

                                highlight line 2 of page 3 in yellow

                                输出：

                                {
                                  "page_3": {
                                    "line_2": {"highlight": "yellow"}
                                  }
                                }

                                输入：

                                red underline the first line on page 2

                                输出：

                                {
                                  "page_2": {
                                    "line_1": {"underline": "red"}
                                  }
                                }
                
                11. 聚合逻辑 (Aggregation Logic)：
                   - 你必须维护一个全局字典对象。
                   - 扫描用户的所有指令，如果多个指令属于同一个 `page_n`，你必须将它们合并。
                   - 严禁在 JSON 顶层出现重复的键。

                12. 强制 JSON 验证步骤 (Self-Correction)：
                   - 在生成最终结果前，检查你的 JSON。如果发现类似 {"page_1": {...}, "page_1": {...}} 的结构，必须将其合并为 {"page_1":        {"line_1":...,"line_4":...}}。
                13. 字体名称规则 (font_name)：
                
                - 当用户提到特定字体或风格时使用。
                - 示例：
                  "改为微软雅黑" -> {"font_name": "Microsoft YaHei"}
                  "使用等宽字体" -> {"font_name": "monospace"}
                  "看起来更正式一点" -> {"font_name": "serif"}
                 14. For 'whole page' or 'all text' requests, use 'line_all' as the key instead of listing lines individually.
                 Example: {'page_1': {'line_all': {'font_name': 'SimHei'}}}"
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
