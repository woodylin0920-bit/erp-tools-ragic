from .le_parser import LEParser
from .tru_parser import TRUParser

# Map folder name (uppercase) → Parser class
PARSERS = {
    "LE":  LEParser,
    "TRU": TRUParser,
}
