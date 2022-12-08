import openpyxl
import json

class ExcelParser:
    #     self.excel_path = excel_path
    #     self.workbook = xlrd.open_workbook(self.excel_path)
    #     self.sheet = self.workbook.sheet_by_index(0)
    #     self.row_num = self.sheet.nrows
    #     self.col_num = self.sheet.ncols

    # def get_data(self):
    #     data = []
    #     for i in range(1, self.row_num):
    #         row_data = self.sheet.row_values(i)
    #         data.append(row_data)
    #     return data

  # TODO how to accept file
    def parse_xlsx_to_json_file(self, excel_path):
      return json.dumps({}, ensure_ascii=False)

excelParser = ExcelParser()
print(excelParser.parse_xlsx_to_json_file("original/test.xlsx"))