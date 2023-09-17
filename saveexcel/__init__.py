from typing import List, Optional
import win32com.client as win32
from pathlib import Path

import xlwt

__all__ = [
            "DEFAULT_OUTPUT_FILENAME",
            "DEFAULT_HEADER_STYLE",
            "DEFAULT_DATA_STYLE",
            "DEFAULT_MAX_ROWS_PER_SHEET",
            "DEFAULT_BASE_SHEET_NAME",
            "EXTRA_SPACE",
            "MAX_CELL_WIDTH",
            "get_adjusted_length",
            "initialize_sheet",
            "convert_xls_to_xlsx"
]

DEFAULT_OUTPUT_FILENAME = "default_output.xls"
DEFAULT_HEADER_STYLE = xlwt.easyxf('font: name 宋体, bold on; align: horiz center, vert center')
DEFAULT_DATA_STYLE = xlwt.easyxf('font: name 微软雅黑; align: horiz center, vert center')
DEFAULT_MAX_ROWS_PER_SHEET = 100
DEFAULT_BASE_SHEET_NAME = "data"
EXTRA_SPACE = 6
MAX_CELL_WIDTH = 255


def initialize_sheet(wb: xlwt.Workbook,
                     sheet_name: str,
                     headers: List[str],
                     header_styling: Optional[xlwt.XFStyle]) -> xlwt.Worksheet:
    """
    Initialize a new sheet with headers.

    Args:
    - wb: The Excel workbook object.
    - sheet_prefix: Prefix for the sheet name.
    - sheet_index: Index to append to the sheet name.
    - headers: List of column headers.
    - header_styling: Style to apply to the headers.

    Returns:
    - sheet: The newly created Excel sheet.
    """
    sheet = wb.add_sheet(sheet_name, cell_overwrite_ok=True)
    for coi_index, header in enumerate(headers):
        adjusted_header_length = get_adjusted_length(header)
        sheet.col(coi_index).width = 257 * (adjusted_header_length + EXTRA_SPACE)
        sheet.write(0, coi_index, header, header_styling)
    return sheet


def get_adjusted_length(cell_str: str) -> int:

    """
    Calculate the adjusted length of a string,
    considering Chinese characters as double the width of English characters.
    """

    return sum(2 if '\u4e00' <= char <= '\u9fff' else 1 for char in cell_str)


def convert_xls_to_xlsx(path: str) -> None:
    path = Path(path)
    excel = win32.gencache.EnsureDispatch('Excel.Application')
    wb = excel.Workbooks.Open(path.absolute())

    # FileFormat=51 is for .xlsx extension
    wb.SaveAs(str(path.absolute().with_suffix(".xlsx")), FileFormat=51)
    wb.Close()
    excel.Application.Quit()


if __name__ == '__main__':
    print(get_adjusted_length("hello"))
    print(get_adjusted_length("你好"))
    print(get_adjusted_length("hello你好"))
