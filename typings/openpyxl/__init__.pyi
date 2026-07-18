from openpyxl import chart as chart
from openpyxl import formatting as formatting
from openpyxl import styles as styles
from openpyxl import utils as utils
from openpyxl import workbook as workbook
from openpyxl import worksheet as worksheet
from openpyxl.workbook.workbook import Workbook as Workbook

def load_workbook(
    filename: str | bytes | object,
    read_only: bool = ...,
    keep_vba: bool = ...,
    data_only: bool = ...,
    keep_links: bool = ...,
    rich_text: bool = ...,
) -> Workbook: ...
