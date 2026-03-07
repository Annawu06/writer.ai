import uno
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
from com.sun.star.style.ParagraphAdjust import CENTER, LEFT, RIGHT, BLOCK

from com.sun.star.ui.dialogs.TemplateDescription import FILEOPEN_SIMPLE

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
    
 

COLOR_MAP = {
    "yellow": 0xFFFF00,
    "red": 0xFF0000,
    "green": 0x00FF00,
    "blue": 0x0000FF,
    "gray": 0xCCCCCC,
    "none": -1
}

def get_doc(ctx):
    smgr = ctx.getServiceManager()
    desktop = smgr.createInstanceWithContext(
        "com.sun.star.frame.Desktop", ctx
    )
    return desktop.getCurrentComponent()
    
    
class Format:

    def __init__(self, ctx):

        self.ctx = ctx
        self.doc = get_doc(ctx)

        if self.doc is None:
            raise RuntimeError("No active document")

        self.controller = self.doc.getCurrentController()
        
    def get_cursor(self):
        """
                获取当前文本光标
        """
        return self.controller.getViewCursor()

    def get_selection(self):
        """
        获取选中的文本对象
        """
        selection = self.controller.getSelection()
        if selection.getCount() > 0:
            return selection.getByIndex(0)
        return None
        
class Format:

    def __init__(self, ctx):

        self.ctx = ctx
        self.doc = get_doc(ctx)

        if self.doc is None:
            raise RuntimeError("No active document")

        self.controller = self.doc.getCurrentController()
        
    def get_cursor(self):
        """
            获取当前文本光标
        """
        return self.controller.getViewCursor()

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

    def set_underline(self,cursor):
        cursor = cursor
        cursor.CharUnderline = SINGLE

    # ------------------------------------------------
    # 字体大小
    # ------------------------------------------------

    def set_font_size(self, cursor,size):
        """
        size: int (例如 12 / 14)
        """
        cursor = cursor
        cursor.CharHeight = size

    # ------------------------------------------------
    # 字体颜色
    # ------------------------------------------------

    def set_font_color(self, cursor, rgb):
        """
        rgb: 例如 0xFF0000
        """
        cursor = cursor
        cursor.CharColor = rgb

    # ------------------------------------------------
    # 高亮
    # ------------------------------------------------
    def highlight(self, cursor, color):
        if isinstance(color, str):
            color = COLOR_MAP.get(color.lower(), 0xFFFF00)

        # 移除这行：view_cursor = self.get_cursor() 
        # 直接使用传入的 cursor。在 LibreOffice 中，TextCursor 已经包含了 Range 信息
        
        try:
            # 确保 color 是整数
            cursor.CharBackColor = int(color)
            log_to_console(f"Highlight applied: {color}")
        except Exception as e:
            log_to_console(f"Highlight failed: {e}")
        

    def remove_highlight(self,cursor):
        cursor = cursor
        cursor.CharBackColor = -1

    # ------------------------------------------------
    # 段落对齐
    # ------------------------------------------------

    def align_center(self,cursor):
        cursor = cursor
        cursor.ParaAdjust = CENTER

    def align_left(self,cursor):
        cursor = cursor
        cursor.ParaAdjust = LEFT

    def align_right(self,cursor):
        cursor = cursor
        cursor.ParaAdjust = RIGHT

    def align_justify(self,cursor):
        cursor = cursor
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

    def clear_format(self,cursor):
        cursor = cursor
        cursor.CharWeight = 100
        cursor.CharPosture = 0
        cursor.CharUnderline = 0
        cursor.CharBackColor = -1
        
        

    
  
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

            line_num = int(line_key.split("_")[1])
            fmt.goto_line(line_num)
            
            cursor = fmt.goto_line(line_num)

            for operation, value in line_value.items():
                if operation in FORMAT_FUNCTION_MAP:
                    func_name = FORMAT_FUNCTION_MAP[operation]
                    func = getattr(fmt, func_name)

                    # 检查 value 是否只是一个开关标记（布尔值或字符串 "true"）
                    is_toggle = value in (True, None) or (isinstance(value, str) and value.lower() == "true")

                    if is_toggle:
                        # 仅传递 cursor，例如调用 set_bold(cursor)
                        func(cursor)
                    else:
                        # 传递 cursor 和具体数值，例如调用 highlight(cursor, "yellow")
                        func(cursor, value)
                        
  
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
        system_prompt = (
            "你是一个专业的文书处理助手，负责将用户的自然语言指令转换为 LibreOffice Writer 的结构化 JSON 数据。"
            "输出必须是一个纯 JSON 字典，不得包含任何解释性文字或 Markdown 代码块标记。"
            "结构参考：{ \"page_n\": { \"line_m\": { \"property\": value } } }。"
            "highlight color must be RGB integer,e.g.:yellow = 16776960"
            "例如输入：'将第一页第一行进行黄色高亮'，输出：{\"page_1\": {\"line_1\": {\"highlight\": \"yellow\"}}}"
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
                # 清理可能存在的 Markdown 标签并解析为字典
                clean_json = content.replace("```json", "").replace("```", "").strip()
                print(f"structured query:{json.loads(clean_json)}")
                return json.loads(clean_json)
            except (json.JSONDecodeError, ValueError) as e:
                print(f"解析 JSON 失败: {e}，原始输出: {content}")
                return {}
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

                file_url = pick_writer_file(self.ctx)

                if not file_url:
                    log_to_console("No file selected.")
                    return

                model = self.desktop.loadComponentFromURL(
                    file_url,
                    "_blank",
                    0,
                    ()
                )

                if not model.supportsService("com.sun.star.text.TextDocument"):
                    log_to_console("Not a Writer document.")
                    return

                user_input = self.input_box(
                    "Document Format:",
                    "AI Formatter",
                    "example: highlight first line"
                )

                if not user_input:
                    return

                format_request = MainJob.askQwen(user_input)

                fmt = Format(self.ctx)

                execute_format_request(format_request, fmt)

                log_to_console("Formatting completed.")

            except Exception as e:

                log_to_console("--- EXCEPTION in trigger(format) ---")
                log_to_console(e)
                traceback.print_exc(file=sys.stderr)

g_ImplementationHelper = unohelper.ImplementationHelper()
g_ImplementationHelper.addImplementation(
    MainJob,
    "org.extension.writerai.do",
    ("com.sun.star.task.Job",),
)
log_to_console("Script loaded, implementation added.")
