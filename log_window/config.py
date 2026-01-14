"""Configuration file for log window"""

from enum import IntEnum, auto, StrEnum

PIPE_TIMEOUT = 3000  # milliseconds


class AppInfo:
    """Application information"""

    PUBLISHER = "OrangeByte"
    PUBLISHER_CREATOR = "Logan Carrozza"
    AUTHOR = f"{PUBLISHER_CREATOR} w/ assistance of GPT/Gemini"
    APP_NAME = "Log Window"
    APP_VERSION = "1.1.0"
    """ Semantic Versioning https://semver.org/
    MAJOR version when you make incompatible API changes
    MINOR version when you adjustment functionality in a backward compatible manner
    PATCH version when you make backward compatible bug fixes
"""


class VerboseLevel(IntEnum):
    """Enumeration representing the level of verbosity for log messages.

    Attributes:
        ERROR: Error messages that indicate a problem or exception.
        INFO: Informational messages about the program's execution.
        WARNING: Messages that indicate a potentially problematic situation.
        DEBUG: Detailed information useful for debugging purposes.
    """

    ERROR = auto()
    INFO = auto()
    WARNING = auto()
    DEBUG = auto()
