"""
Configuration file for the Batch Render extension.

This file contains variables that are used to store application information, such as the
publisher, creator, author, and application version.
"""


class Config:
    """Contains application information."""
    PUBLISHER = "OrangeByte"
    PUBLISHER_CREATOR = "Logan Carrozza"
    AUTHOR = f"{PUBLISHER_CREATOR} w/ assistance of GPT"
    APP_NAME = "Vray Batch Render"

    APP_VERSION = "1.0.6"
    """ Semantic Versioning https://semver.org/
    MAJOR version when you make incompatible API changes
    MINOR version when you adjustment functionality in a backward compatible manner
    PATCH version when you make backward compatible bug fixes
    """

    BATCH_RENDER_SETTINGS = "batchRenderSettings"  # Can't contain spaces
    DEFAULT_TEXT = "-" * 80
    DEFAULT_PATH_TEXT = "Default Path + Name"
    UUID_PARAMETER_NAME = "PersistentID"
