
import xlwt
from typing import List, Optional, Any
import json
from saveexcel import *
import os


def save_to_excel(data_list: List[List[Any]],
                  column_headers: List[str],
                  header_style: Optional[xlwt.XFStyle] = DEFAULT_HEADER_STYLE,
                  data_style: Optional[xlwt.XFStyle] = DEFAULT_DATA_STYLE,
                  output_filename: str = DEFAULT_OUTPUT_FILENAME,
                  max_rows_per_sheet: int = DEFAULT_MAX_ROWS_PER_SHEET,
                  base_sheet_name: str = DEFAULT_BASE_SHEET_NAME, ) -> None:
    """
    Save data to an Excel file with specified headers and styles.

    Args:
    - data_list: List of data rows to be saved.
    - output_filename: Name of the output Excel file.
    - column_headers: List of column headers.
    - header_style: Style for the headers.
    - data_style: Style for the data cells.
    - max_rows_per_sheet: Maximum number of rows per sheet.
    - base_sheet_name: Base name for the sheets.

    Returns:
    None
    """

    # Initialize workbook and settings
    main_workbook = xlwt.Workbook(encoding="utf-8")
    current_sheet = initialize_sheet(main_workbook, f"{base_sheet_name}_{1}", column_headers, header_style)
    max_column_lengths = [get_adjusted_length(header) for header in column_headers]

    sheet_count = 1
    current_row_index = 1
    count = 0
    # Write data to the Excel sheets
    for data_row in data_list:

        if current_row_index == max_rows_per_sheet:
            sheet_count += 1
            current_sheet = initialize_sheet(main_workbook, f"{base_sheet_name}_{sheet_count}",
                                             column_headers, header_style)
            current_row_index = 1
            max_column_lengths = [get_adjusted_length(header) for header in column_headers]

        for idx, cell_value in enumerate(data_row):
            cell = str(cell_value)

            # Adjust column width if current data length is bigger
            adjusted_length = get_adjusted_length(cell)
            if adjusted_length > max_column_lengths[idx]:
                max_column_lengths[idx] = adjusted_length
                calculated_width = 257 * (max_column_lengths[idx] + EXTRA_SPACE)
                # Ensure the width does not exceed the maximum allowable width
                current_sheet.col(idx).width = min(calculated_width, 65535)
            current_sheet.write(current_row_index, idx, cell_value, data_style)
        count += 1
        current_row_index += 1
        print(f"{'=' * 30} Data item number {count} saved successfully {'=' * 30}")

    # Save the workbook and display a success message

    output_filenames = output_filename[:-1] if output_filename.endswith(".xlsx") else output_filename
    main_workbook.save(output_filenames)
    if output_filename.endswith(".xlsx"):
        convert_xls_to_xlsx(output_filenames)
        os.remove(output_filenames)
    print(f"{'=' * 30} Total {len(data_list)} records saved successfully {'=' * 30}")


class ExcelWriter:
    def __init__(self,
                 column_headers: List[str],
                 header_style: Optional[xlwt.XFStyle] = DEFAULT_HEADER_STYLE,
                 data_style: Optional[xlwt.XFStyle] = DEFAULT_DATA_STYLE,
                 max_rows_per_sheet: int = DEFAULT_MAX_ROWS_PER_SHEET,
                 output_filename: str = DEFAULT_OUTPUT_FILENAME,
                 base_sheet_name: str = DEFAULT_BASE_SHEET_NAME):
        self.output_filename = output_filename
        self.column_headers = column_headers
        self.header_style = header_style
        self.data_style = data_style
        self.max_rows_per_sheet = max_rows_per_sheet
        self.base_sheet_name = base_sheet_name
        self.workbook = xlwt.Workbook(encoding="utf-8")
        self.current_sheet = None
        self.sheet_count = 1
        self.current_row_index = 1
        self.count = 0
        self.max_column_lengths = [get_adjusted_length(header) for header in column_headers]

    def _initialize_sheet(self) -> None:
        self.current_sheet = self.workbook.add_sheet(f"{self.base_sheet_name}_{self.sheet_count}",
                                                     cell_overwrite_ok=True)
        for coi_index, header in enumerate(self.column_headers):
            adjusted_header_length = get_adjusted_length(header)
            self.current_sheet.col(coi_index).width = 257 * (adjusted_header_length + EXTRA_SPACE)
            self.current_sheet.write(0, coi_index, header, self.header_style)

    def save_data(self, data_list: List[List[Any]]) -> None:
        self._initialize_sheet()

        for data_row in data_list:
            if self.current_row_index == self.max_rows_per_sheet:
                self.sheet_count += 1
                self._initialize_sheet()
                self.current_row_index = 1
                self.max_column_lengths = [get_adjusted_length(header) for header in self.column_headers]

            for idx, cell_value in enumerate(data_row):
                cell = str(cell_value)
                adjusted_length = get_adjusted_length(cell)
                if adjusted_length > self.max_column_lengths[idx]:
                    self.max_column_lengths[idx] = adjusted_length
                    calculated_width = 257 * (self.max_column_lengths[idx] + EXTRA_SPACE)
                    self.current_sheet.col(idx).width = min(calculated_width, 65535)
                self.current_sheet.write(self.current_row_index, idx, cell_value, self.data_style)
            self.count += 1
            self.current_row_index += 1
            print(f"{'=' * 30} Data item number {self.count} written successfully {'=' * 30}")
        output_filename = self.output_filename[:-1] if self.output_filename.endswith(".xlsx") else self.output_filename
        self.workbook.save(output_filename)
        if self.output_filename.endswith(".xlsx"):
            convert_xls_to_xlsx(output_filename)
            os.remove(output_filename)
        print(f"{'=' * 30} Total {len(data_list)} records saved successfully {'=' * 30}")


# Function usage:
# if __name__ == '__main__':
#     with open('sample_items.json', "r", encoding="utf-8") as file:
#         p_data_list = json.load(file)
#     data_list1 = [list(item.values()) for item in p_data_list]
#     p_headers = [item for item in p_data_list[0]]
#     save_to_excel(data_list1, p_headers, output_filename="sample_itemws.xls")

# Class usage:
if __name__ == '__main__':
    with open('sample_items.json', "r", encoding="utf-8") as file:
        p_data_list = json.load(file)
    data_list2 = [list(item.values()) for item in p_data_list]
    p_headers = [item for item in p_data_list[0]]
    excel_writer = ExcelWriter(p_headers, output_filename="sample_items1.xlsx", max_rows_per_sheet=5000)
    excel_writer.save_data(data_list2)
