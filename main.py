from batch_render import BatchRenderDialog
from log_window import Console

try:
    console = Console()
    main_dialog = BatchRenderDialog()
    main_dialog.exec()
except Exception as main_error:
    raise main_error
finally:
    # noinspection PyUnboundLocalVariable
    console.shutdown()
