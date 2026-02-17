import uno
import unohelper
import json
import os
import logging
import traceback
from com.sun.star.task import XJobExecutor
from com.sun.star.awt import MessageBoxButtons as MSG_BUTTONS
from com.sun.star.awt import XActionListener, XItemListener
from com.sun.star.awt.PosSize import POS, SIZE, POSSIZE
from com.sun.star.awt.PushButtonType import OK, CANCEL
from com.sun.star.util.MeasureUnit import TWIP

_debug_logging_enabled = False

def log_to_file(message):
    if not _debug_logging_enabled:
        return
    log_dir = os.path.join(os.path.expanduser('~'), '.writerai')
    os.makedirs(log_dir, exist_ok=True)
    log_file_path = os.path.join(log_dir, 'log.txt')
    logging.basicConfig(filename=log_file_path, level=logging.INFO, format='%(asctime)s - %(message)s')
    logging.info(message)

class MainJob(unohelper.Base, XJobExecutor):
    def __init__(self, ctx):
        self.ctx = ctx
        try:
            self.sm = ctx.getServiceManager()
            self.desktop = XSCRIPTCONTEXT.getDesktop()
        except NameError:
            self.sm = ctx.ServiceManager
            self.desktop = self.ctx.getServiceManager().createInstanceWithContext(
                "com.sun.star.frame.Desktop", self.ctx)

    def show_message_box(self, title, message):
        try:
            print("show message box")
            frame = self.desktop.getCurrentFrame()
            window = frame.getContainerWindow()
            toolkit = window.getToolkit()
            msgbox = toolkit.createMessageBox(window, MSG_BUTTONS.BUTTONS_OK, title, str(message))
            msgbox.execute()
        except Exception as e:
            log_to_file(f"Failed to show message box: {e}\n{traceback.format_exc()}")

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
            log_to_file(f"Error writing to {config_file_path}: {e}")

    def _as_bool(self, value):
        if isinstance(value, str):
            return value.lower() in ('true', '1', 't', 'y', 'yes')
        return bool(value)

    BACKEND_PRESETS = [
        ("Gemini 3.5 Pro", "chat", "https://generativelanguage.googleapis.com/v1beta"),
        ("Gemini 3.5 Fast", "chat", "https://generativelanguage.googleapis.com/v1beta"),
        ("QWen", "chat", "https://dashscope.aliyuncs.com/api/v1/services/aigc/text-generation/generation"),
    ]

    def _detect_backend(self):
        model_name = self.get_config("model", "").lower()
        for i, preset in enumerate(self.BACKEND_PRESETS):
            if preset[0].lower() == model_name:
                return i
        return 0

    def _read_dialog_config(self, controls):
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
        if not result:
            return
        for key, value in result.items():
            self.set_config(key, value)

    def settings_box(self, title="", x=None, y=None):
        WIDTH = 600
        HEIGHT = 150 # Increased height to ensure buttons are visible
        HORI_MARGIN = 10
        VERT_MARGIN = 10
        BUTTON_WIDTH = 100
        BUTTON_HEIGHT = 26
        HORI_SEP = 10
        VERT_SEP = 10
        LABEL_WIDTH = 150
        EDIT_HEIGHT = 24

        ctx = self.ctx
        def create(name):
            return ctx.getServiceManager().createInstanceWithContext(name, ctx)

        try:
            dialog = create("com.sun.star.awt.UnoControlDialog")
            dialog_model = create("com.sun.star.awt.UnoControlDialogModel")
            dialog.setModel(dialog_model)
            dialog.setTitle(title)
            dialog.setPosSize(0, 0, WIDTH, HEIGHT, SIZE)

            def add(name, ctrl_type, x, y, width, height, props):
                model = dialog_model.createInstance("com.sun.star.awt.UnoControl" + ctrl_type + "Model")
                dialog_model.insertByName(name, model)
                control = dialog.getControl(name)
                control.setPosSize(x, y, width, height, POSSIZE)
                for key, value in props.items():
                    setattr(model, key, value)
                return control

            controls = {}
            edit_width = WIDTH - HORI_MARGIN * 2 - LABEL_WIDTH - HORI_SEP
            
            # --- Model Preset ---
            y_pos = VERT_MARGIN
            add("label_backend", "FixedText", HORI_MARGIN, y_pos + 4, LABEL_WIDTH, LABEL_HEIGHT, {"Label": "Model Preset:"})
            backend_names = tuple(p[0] for p in self.BACKEND_PRESETS)
            current_backend_idx = self._detect_backend()
            controls["backend"] = add("list_backend", "ListBox", HORI_MARGIN + LABEL_WIDTH, y_pos,
                edit_width, EDIT_HEIGHT,
                {"Dropdown": True, "StringItemList": backend_names, "SelectedItems": (current_backend_idx,), "LineCount": len(self.BACKEND_PRESETS)})
            
            # --- API Key ---
            y_pos += EDIT_HEIGHT + VERT_SEP
            add("label_api_key", "FixedText", HORI_MARGIN, y_pos + 4, LABEL_WIDTH, LABEL_HEIGHT, {"Label": "API Key:"})
            controls["api_key"] = add("edit_api_key", "Edit", HORI_MARGIN + LABEL_WIDTH, y_pos,
                edit_width, EDIT_HEIGHT, {"Text": str(self.get_config("api_key", "")), "PasswordChar": "*"})

            # --- Buttons ---
            y_pos = HEIGHT - BUTTON_HEIGHT - VERT_MARGIN
            button_start_x = (WIDTH - (BUTTON_WIDTH * 2 + HORI_SEP)) / 2
            add("btn_ok", "Button", button_start_x, y_pos, BUTTON_WIDTH, BUTTON_HEIGHT, {"PushButtonType": OK, "DefaultButton": True})
            add("btn_cancel", "Button", button_start_x + BUTTON_WIDTH + HORI_SEP, y_pos, BUTTON_WIDTH, BUTTON_HEIGHT, {"PushButtonType": CANCEL})

            # --- Create Peer and Execute ---
            frame = self.desktop.getCurrentFrame()
            window = frame.getContainerWindow() if frame else None
            if not window:
                self.show_message_box("Error", "Could not get window to create dialog.")
                return {}
            
            dialog.createPeer(create("com.sun.star.awt.Toolkit"), window)
            
            ret = {}
            if dialog.execute():
                ret = self._read_dialog_config(controls)
            
        except Exception as e:
            self.show_message_box("Critical Error", f"Failed to create settings_box:\n{e}\n{traceback.format_exc()}")
            ret = {}
        finally:
            if 'dialog' in locals() and dialog:
                dialog.dispose()
        
        return ret

    def input_box(self, message, title="", default="", x=None, y=None):
        WIDTH = 500
        HORI_MARGIN = 10
        VERT_MARGIN = 10
        BUTTON_WIDTH = 80
        BUTTON_HEIGHT = 25
        HORI_SEP = 10
        VERT_SEP = 10
        LABEL_WIDTH = 100
        EDIT_HEIGHT = 25 * 5

        HEIGHT = VERT_MARGIN * 3 + EDIT_HEIGHT + BUTTON_HEIGHT

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

        add("label", "FixedText", HORI_MARGIN, VERT_MARGIN + 5, LABEL_WIDTH, 25, {"Label": str(message)})
        
        edit_x = HORI_MARGIN + LABEL_WIDTH + HORI_SEP
        edit_width = WIDTH - edit_x - HORI_MARGIN
        add("edit", "Edit", edit_x, VERT_MARGIN, edit_width, EDIT_HEIGHT, 
            {"Text": str(default), "MultiLine": True, "VScroll": True})
        
        buttons_y = VERT_MARGIN + EDIT_HEIGHT + VERT_SEP

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
        global _debug_logging_enabled
        _debug_logging_enabled = self._as_bool(self.get_config("debug_logging", False))

        model = self.desktop.getCurrentComponent()

        if hasattr(model, "Text"):
            if args == "format":
                user_input = self.input_box("Input format:", "AI Formatter", "example:highlight the first line on page 1")
                if user_input:
                    text = model.Text
                    cursor = model.getCurrentController().getViewCursor()
                    text.insertString(cursor, f"User entered: {user_input}", 0)
            elif args == "setting":
                try:
                    result = self.settings_box("Writer.ai Settings")
                    self._save_settings(result)
                except Exception as e:
                    error_message = f"An error occurred in the settings dialog:\n\n{e}\n\n{traceback.format_exc()}"
                    log_to_file(error_message)
                    self.show_message_box("Settings Error", error_message)

g_ImplementationHelper = unohelper.ImplementationHelper()
g_ImplementationHelper.addImplementation(
    MainJob,
    "org.extension.writerai.do",
    ("com.sun.star.task.Job",),
)
