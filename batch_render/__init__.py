"""
V-Ray Batch Render for 3ds Max

This module provides an enhanced batch render dialog for 3ds Max with V-Ray integration.
It extends the standard batch render functionality with features like multi-row editing,
render queue management, and V-Ray specific settings.

Example usage:
    from batch_render import BatchRenderDialog

    # Create and show the batch render dialog
    dialog = BatchRenderDialog()
    dialog.show()
"""

from .main import BatchRenderDialog
from .config import Config

__version__ = Config.APP_VERSION
__author__ = Config.AUTHOR

__all__ = ["BatchRenderDialog", "Config"]
