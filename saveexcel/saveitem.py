import json
from typing import List, Optional, Union
import xlwt
import xlrd
from xlutils.copy import copy

from saveexcel import *


def save_to_excel(
                  data_item: Union[List[str], List[List[str]]],
                  column_headers: List[str],
                  output_filename: str = DEFAULT_OUTPUT_FILENAME,
                  header_style: Optional[xlwt.XFStyle] = DEFAULT_HEADER_STYLE,
                  data_style: Optional[xlwt.XFStyle] = DEFAULT_DATA_STYLE,
                  max_rows_per_sheet: int = DEFAULT_MAX_ROWS_PER_SHEET,
                  base_sheet_name: str = DEFAULT_BASE_SHEET_NAME,
                  extra_space: int = EXTRA_SPACE
                  ) -> None:

    """Save data to an Excel file, potentially across multiple sheets."""

    # Ensure data_items is a list of lists
    if not isinstance(data_item[0], list):
        data_items = [data_item]
    else:
        data_items = data_item

    # Check if the file already exists
    try:
        rb = xlrd.open_workbook(output_filename, formatting_info=True)
        last_sheet_rb = rb.sheet_by_index(-1)
        current_row_index = last_sheet_rb.nrows
        sheet_count = rb.nsheets

        main_workbook = copy(rb)

    except FileNotFoundError:
        main_workbook = xlwt.Workbook(encoding="utf-8")
        initialize_sheet(main_workbook, f"{base_sheet_name}_1", column_headers, header_style)
        current_row_index = 1
        sheet_count = 1

    # Write data to the Excel sheet
    for item in data_items:

        if current_row_index == max_rows_per_sheet:
            sheet_count += 1
            initialize_sheet(main_workbook, f"{base_sheet_name}_{sheet_count}", column_headers, header_style)
            current_row_index = 1

        current_sheet = main_workbook.get_sheet(-1)

        for idx, cell_value in enumerate(item):
            adjusted_length = get_adjusted_length(str(cell_value))
            calculated_width = 257 * min(adjusted_length + extra_space, MAX_CELL_WIDTH)
            if current_sheet.col(idx).width < calculated_width:
                current_sheet.col(idx).width = min(calculated_width, MAX_CELL_WIDTH * 257)
            current_sheet.write(current_row_index, idx, cell_value, data_style)

        current_row_index += 1
        count = (max_rows_per_sheet - 1) * (sheet_count - 1) + current_row_index - 1
        print(f"{'=' * 30} Data item number {count} saved successfully {'=' * 30}")

    # Save the workbook
    main_workbook.save(output_filename)


class ExcelSaver:

    """
    Class to save data to Excel in an object-oriented manner.
    """
    def __init__(self,
                 column_headers: List[str],
                 header_style: Optional[xlwt.XFStyle] = DEFAULT_HEADER_STYLE,
                 data_style: Optional[xlwt.XFStyle] = DEFAULT_DATA_STYLE,
                 output_filename: str = DEFAULT_OUTPUT_FILENAME,
                 max_rows_per_sheet: int = DEFAULT_MAX_ROWS_PER_SHEET,
                 base_sheet_name: str = DEFAULT_BASE_SHEET_NAME,
                 extra_space: int = EXTRA_SPACE):
        self.output_filename = output_filename
        self.column_headers = column_headers
        self.header_style = header_style
        self.data_style = data_style
        self.max_rows_per_sheet = max_rows_per_sheet
        self.base_sheet_name = base_sheet_name
        self.item_count = 0
        self.extra_space = extra_space

    def initialize_sheet(self, workbook, sheet_name):
        """Initialize a new sheet and set headers."""
        sheet = workbook.add_sheet(sheet_name, cell_overwrite_ok=True)
        for idx, header in enumerate(self.column_headers):
            adjusted_length = get_adjusted_length(header)
            sheet.col(idx).width = 257 * min(adjusted_length + self.extra_space, MAX_CELL_WIDTH)
            sheet.write(0, idx, header, self.header_style)
        return sheet

    def save_data_item(self, data_item: Union[List[str], List[List[str]]]) -> None:
        """Save a single data item to the Excel file."""
        try:
            rb = xlrd.open_workbook(self.output_filename, formatting_info=True)
            main_workbook = copy(rb)
            last_sheet_rb = rb.sheet_by_index(-1)
            current_row_index = last_sheet_rb.nrows
            main_workbook.get_sheet(-1)
            sheet_count = rb.nsheets
        except FileNotFoundError:
            main_workbook = xlwt.Workbook(encoding="utf-8")
            self.initialize_sheet(main_workbook, f"{self.base_sheet_name}_1")
            current_row_index = 1
            sheet_count = 1

        # Ensure data_items is a list of lists
        if not isinstance(data_item[0], list):
            data_items = [data_item]
        else:
            data_items = data_item

        for item in data_items:

            if current_row_index >= self.max_rows_per_sheet:
                self.initialize_sheet(main_workbook, f"{self.base_sheet_name}_{sheet_count + 1}")
                current_row_index = 1
                sheet_count += 1

            current_sheet = main_workbook.get_sheet(-1)

            for idx, cell_value in enumerate(item):
                adjusted_length = get_adjusted_length(str(cell_value))
                calculated_width = 257 * min(adjusted_length + self.extra_space, MAX_CELL_WIDTH)
                if current_sheet.col(idx).width < calculated_width:
                    current_sheet.col(idx).width = min(calculated_width, MAX_CELL_WIDTH * 257)
                current_sheet.write(current_row_index, idx, cell_value, self.data_style)

            current_row_index += 1
            count = (self.max_rows_per_sheet - 1) * (sheet_count - 1) + current_row_index - 1
            print(f"{'=' * 30} Data item number {count} saved successfully {'=' * 30}")
        main_workbook.save(self.output_filename)


# Function usage:
# if __name__ == '__main__':
#     with open('sample_items.json', "r", encoding="utf-8") as file:
#         p_data_list = json.load(file)
#     data_list1 = [list(item.values()) for item in p_data_list]
#     p_headers = [item for item in p_data_list[0]]
#     save_to_excel(data_list1[:501], p_headers, output_filename="sample_itemws.xls")


# Class usage:
if __name__ == '__main__':
    with open('sample_items.json', "r", encoding="utf-8") as file:
        p_data_list = json.load(file)
    data_list1 = [list(item.values()) for item in p_data_list]
    p_headers = [item for item in p_data_list[0]]
    excel_saver = ExcelSaver(p_headers, output_filename="sample_itemws.xls")
    for item in data_list1[:1000]:
        excel_saver.save_data_item(item)
    #
    # for item in data_list1[:1000]:
    # excel_saver.save_data_item(data_list1[:1000])
