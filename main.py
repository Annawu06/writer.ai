import uno
import unohelper
from com.sun.star.task import XJobExecutor

class MainJob(unohelper.Base, XJobExecutor):
    def __init__(self, ctx):
        self.ctx = ctx
        try:
            self.sm = ctx.getServiceManager()
            self.desktop = XSCRIPTCONTEXT.getDesktop()
            self.document = XSCRIPTCONTEXT.getDocument()
        except NameError:
            self.sm = ctx.ServiceManager
            self.desktop = self.ctx.getServiceManager().createInstanceWithContext(
                "com.sun.star.frame.Desktop", self.ctx)

    def input_box(self, message, title="", default="", x=None, y=None):
        """ Shows dialog with input box. """
        WIDTH = 400
        HORI_MARGIN = 10
        VERT_MARGIN = 10
        BUTTON_WIDTH = 80
        BUTTON_HEIGHT = 25
        HORI_SEP = 10
        VERT_SEP = 10
        LABEL_WIDTH = 100
        EDIT_HEIGHT = 25

        # Calculate total height
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

        # Label
        add("label", "FixedText", HORI_MARGIN, VERT_MARGIN + 5, LABEL_WIDTH, EDIT_HEIGHT, {"Label": str(message)})
        
        # Edit field
        edit_x = HORI_MARGIN + LABEL_WIDTH + HORI_SEP
        edit_width = WIDTH - edit_x - HORI_MARGIN
        add("edit", "Edit", edit_x, VERT_MARGIN, edit_width, EDIT_HEIGHT, {"Text": str(default)})
        
        # Buttons Y position
        buttons_y = VERT_MARGIN + EDIT_HEIGHT + VERT_SEP

        # OK Button
        ok_x = (WIDTH - (BUTTON_WIDTH * 2 + HORI_SEP)) / 2
        add("btn_ok", "Button", ok_x, buttons_y, BUTTON_WIDTH, BUTTON_HEIGHT, {"PushButtonType": OK, "DefaultButton": True})
        
        # Cancel Button
        cancel_x = ok_x + BUTTON_WIDTH + HORI_SEP
        add("btn_cancel", "Button", cancel_x, buttons_y, BUTTON_WIDTH, BUTTON_HEIGHT, {"PushButtonType": CANCEL})

        frame = self.desktop.getCurrentFrame()
        window = frame.getContainerWindow() if frame else None
        dialog.createPeer(create("com.sun.star.awt.Toolkit"), window)
        
        if x is not None and y is not None:
            ps = dialog.convertSizeToPixel(uno.createUnoStruct("com.sun.star.awt.Size", x, y), TWIP)
            _x, _y = ps.Width, ps.Height
        elif window:
            ps = window.getPosSize()
            _x = ps.Width / 2 - WIDTH / 2
            _y = ps.Height / 2 - HEIGHT / 2
        dialog.setPosSize(_x, _y, 0, 0, POS)
        
        edit = dialog.getControl("edit")
        edit.setSelection(uno.createUnoStruct("com.sun.star.awt.Selection", 0, len(str(default))))
        edit.setFocus()
        
        ret = edit.getModel().Text if dialog.execute() else ""
        dialog.dispose()
        return ret

    def trigger(self, args):
        model = self.desktop.getCurrentComponent()

        if hasattr(model, "Text"):
            if args == "format":
                user_input = self.input_box("Input format:", "AI Formatter", "example:highlight the first line on page 1")
                if user_input:
                    text = model.Text
                    cursor = model.getCurrentController().getViewCursor()
                    text.insertString(cursor, f"User entered: {user_input}", 0)
            elif args == "setting":
                text = model.Text
                cursor = model.getCurrentController().getViewCursor()
                text.insertString(cursor, "Setting action triggered!", 0)

g_ImplementationHelper = unohelper.ImplementationHelper()
g_ImplementationHelper.addImplementation(
    MainJob,
    "org.extension.writerai.do",
    ("com.sun.star.task.Job",),
)
