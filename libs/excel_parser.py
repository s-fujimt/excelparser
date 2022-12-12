import openpyxl
import json
from libs.color_helper import theme_and_tint_to_rgb

class ExcelParser:
    workbook = None
    current_sheet = None
    current_sheet_first_empty_row = None
    current_range = None

    def __get_merged_ranges(self):
        return self.current_sheet.merged_cells.ranges

    def __is_merged_cell(self, cell, merged_ranges):
        for merged_range in merged_ranges:
            if cell.coordinate in merged_range:
                self.current_range = {
                    "columns": merged_range.size["columns"],
                    "rows": merged_range.size["rows"]
                }
                return True
        return False

    def __get_merged_cell_data(self):
        cell_data = {}
        if self.current_range["columns"] > 1:
            cell_data["colspan"] = self.current_range["columns"]
        if self.current_range["rows"] > 1:
            cell_data["rowspan"] = self.current_range["rows"]
        return cell_data

    def __set_first_empty_row(self):
        self.current_sheet_first_empty_row = self.current_sheet.max_row + 1

    def __get_default_font_data(self):
        return {
          "font": self.current_sheet.cell(row=self.current_sheet_first_empty_row, column=1).font.name,
          "size": int(self.current_sheet.cell(row=self.current_sheet_first_empty_row, column=1).font.size)
        }
    
    def __get_cell_alignment(self, cell):
        alignment = {}
        if cell.alignment:
          if cell.alignment.horizontal in ["center", "left", "right"]:
            alignment["horizontal"] = cell.alignment.horizontal
          if cell.alignment.vertical in ["center", "bottom", "top"]:
            alignment["vertical"] = cell.alignment.vertical
        return alignment

    def __get_color_from_theme(self, color_data):
        color = theme_and_tint_to_rgb(self.workbook, color_data.theme, color_data.tint)
        return color

    def __get_color_data(self, color_data):
        color = None
        if color_data.type == "rgb":
            if color_data.rgb == "00000000":
                color = "FFFFFF"
            else:
              color = color_data.rgb[2:]
        if color_data.type == "indexed":
            if color_data.indexed == 63:
                color = "FFFFFF"
            if color_data.indexed == 64:
                color = "000000"
            else:
              Colors = openpyxl.styles.colors.COLOR_INDEX
              color = Colors[color_data.indexed][2:]
        if color_data.type == "theme":
            color = self.__get_color_from_theme(color_data)
        return f"#{color}"

    def __get_cell_font_data(self, cell):
        cell_font_data = {}
        default_font_data = self.__get_default_font_data()
        if cell.font:
          if cell.font.name != default_font_data["font"]:
            cell_font_data["font"] = cell.font.name
          if cell.font.size and cell.font.size != default_font_data["size"]:
            cell_font_data["size"] = int(cell.font.size)
          if cell.font.bold:
            cell_font_data["style"] = "bold"
          if cell.font.underline:
            cell_font_data["underline"] = cell.font.underline
          if cell.font.strikethrough:
            cell_font_data["strikethrough"] = cell.font.strikethrough
          if cell.font.color:
            color = self.__get_color_data(cell.font.color)
            if color and color != "#000000":
              cell_font_data["color"] = color

        return cell_font_data

    def __get_border_style(self, border_style):
        if border_style == "medium":
            return "thick"
        if border_style == "thick":
            return "extrathick"
        if border_style == "double":
            return "double"
        else:
          return "single"
    
    def __get_cell_border_data(self, cell, is_merged_cell):
        cell_border_data = {}
        outline = {}

        border = cell.border

        if is_merged_cell:
            border_style = border.top.style
            if border_style:
                outline["style"] = self.__get_border_style(border_style)

            border_color = border.top.color
            if border_color:
                outline["color"] = self.__get_color_data(border_color)
        else:
            if border.top and border.right and border.bottom and border.left:
                border = [border.top, border.right, border.bottom, border.left]

                border_style = list(set([b.style for b in border]))
                border_color = list(set([b.color for b in border]))

                if len(border_style) > 1 and None in border_style:
                    border_style.remove(None)
                    border_color.remove(None)

                if len(border_style) == 1 and len(border_color) == 1:
                    if border_style[0]:
                        outline["style"] = self.__get_border_style(border_style[0])
                    if border_color[0]:
                        outline["color"] = self.__get_color_data(border_color[0])
                else:
                    for side in ["top", "right", "bottom", "left"]:
                        if getattr(cell.border, side).style:
                            cell_border_data[side] = {
                                "style": self.__get_border_style(getattr(cell.border, side).style),
                                "color": self.__get_color_data(getattr(cell.border, side).color)
                            }

        if outline:
            cell_border_data["outline"] = outline
    
        return cell_border_data

    def __get_fill_color(self, cell):
        if cell.fill:
            color = self.__get_color_data(cell.fill.start_color)
            if color and color != "#FFFFFF":
                return color
        return None

    def __get_cell_data(self, cell):
        value = cell.value

        if value is None:
            return None

        cell_data = {
            "value": value,
            "colnumber": cell.coordinate[0]
        }

        is_merged_cell = self.__is_merged_cell(cell, self.__get_merged_ranges())
        if is_merged_cell:
            cell_data.update(self.__get_merged_cell_data())

        alignment = self.__get_cell_alignment(cell)
        if alignment:
          cell_data["alignment"] = alignment
        
        font = self.__get_cell_font_data(cell)
        if font:
          cell_data["font"] = font if font else None

        border = self.__get_cell_border_data(cell, is_merged_cell)
        if border:
          cell_data["border"] = border

        fill_color = self.__get_fill_color(cell)
        if fill_color:
          cell_data["fill"] = {"color":fill_color}

        return cell_data

    def __get_row_data(self, row):
        row_data = {
            "linenumber": row[0].row
        }

        columns = []
        for cell in row:
            cell_data = self.__get_cell_data(cell)
            if cell_data:
                columns.append(cell_data)
        
        if columns:
            row_data["columns"] = columns
        return row_data

    def __get_rows(self):
        rows = []
        for row in self.current_sheet.iter_rows():
            if row[0].row == self.current_sheet_first_empty_row:
                break
            row_data = self.__get_row_data(row)
            if row_data:
              rows.append(row_data)
        return rows


    def __get_sheet_data(self, sheetnumber):
        font_data = self.__get_default_font_data()
        rows = self.__get_rows()

        sheet_data = {
            "sheetnumber": sheetnumber,
            "sheetname": self.current_sheet.title,
            "font": font_data,
            "lines": rows
        }

        return sheet_data

    def __open_workbook(self, excel_path):
      workbook = openpyxl.load_workbook(excel_path)
      return workbook

    def parse_xlsx_to_json_file(self, excel_path):
        self.workbook = self.__open_workbook(excel_path)
        sheet_names = self.workbook.sheetnames

        sheet_data = []
        for i, sheetname in enumerate(sheet_names):
            self.current_sheet = self.workbook[sheetname]
            self.__set_first_empty_row()
            sheet_data.append(self.__get_sheet_data(i+1))

        return json.dumps({"sheets": sheet_data}, ensure_ascii=False)
