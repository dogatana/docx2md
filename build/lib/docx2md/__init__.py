from .docxfile import DocxFile, DocxFileError
from .converter import Converter
from .docxmedia import DocxMedia

VERSION = (1, 0, 0)

__version__ = ".".join(map(str, VERSION))
