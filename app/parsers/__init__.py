from .le_parser import LEParser
from .tru_parser import TRUParser
from .template_parser import TemplateParser

# Map folder name (uppercase) → Parser class
PARSERS = {
    "LE":       LEParser,
    "TRU":      TRUParser,
    "TEMPLATE": TemplateParser,
}
