import win32com.client as win32
from pathlib import Path


def convert_xls_to_xlsx(path: str) -> None:
    path = Path(path)
    excel = win32.gencache.EnsureDispatch('Excel.Application')
    wb = excel.Workbooks.Open(path.absolute())

    # FileFormat=51 is for .xlsx extension
    wb.SaveAs(str(path.absolute().with_suffix(".xls")), FileFormat=51)
    wb.Close()
    excel.Application.Quit()


def convert_xlsx_to_xls(path: str) -> None:
    path = Path(path)
    excel = win32.gencache.EnsureDispatch('Excel.Application')
    wb = excel.Workbooks.Open(path.absolute())

    # FileFormat=51 is for .xlsx extension
    wb.SaveAs(str(path.absolute().with_suffix(".xls")), FileFormat=51)
    wb.Close()
    excel.Application.Quit()


# Usage
if __name__ == '__main__':
    convert_xls_to_xlsx('sample_itemws.xls')
