import openpyxl
import json
from libs.color_helper import theme_and_tint_to_rgb

class ExcelParser:
    workbook = None
    current_sheet = None
    current_sheet_first_empty_row = None
    current_sheet_first_empty_column = None
    current_range = None

    def __get_merged_ranges(self):
        print("Getting merged ranges")
        return self.current_sheet.merged_cells.ranges if self.current_sheet.merged_cells else []

    def __is_merged_cell(self, cell, merged_ranges):
        print("Checking if cell is merged")
        for merged_range in merged_ranges:
            if cell.coordinate in merged_range:
                self.current_range = {
                    "columns": merged_range.size["columns"],
                    "rows": merged_range.size["rows"]
                }
                return True
        return False

    def __get_merged_cell_data(self):
        print("Getting merged cell data")
        try:
            cell_data = {}
            if self.current_range["columns"] > 1:
                cell_data["colspan"] = self.current_range["columns"]
            if self.current_range["rows"] > 1:
                cell_data["rowspan"] = self.current_range["rows"]
            return cell_data
        except:
            print("Error getting merged cell data")
            return {}

    def __set_first_empty_row(self):
        print("Setting first empty row")
        try:
            for row in range(self.current_sheet.max_row, 0, -1):
                for column in range(1, self.current_sheet.max_column + 1):
                    cell = self.current_sheet.cell(row=row, column=column)
                    if cell.value or self.__has_fill_color(cell) or self.__has_border(cell):
                        self.current_sheet_first_empty_row = row + 1
                        return
            if not self.current_sheet_first_empty_row:
                self.current_sheet_first_empty_row = 1
        except:
            print("Error getting first empty row")


    def __set_first_empty_column(self):
        print("Setting first empty column")
        try:
            for column in range(self.current_sheet.max_column, 0, -1):
                for row in range(1, self.current_sheet.max_row + 1):
                    cell = self.current_sheet.cell(row=row, column=column)
                    if cell.value or self.__has_fill_color(cell) or self.__has_border(cell):
                        self.current_sheet_first_empty_column = column + 1
                        return
            if not self.current_sheet_first_empty_column:
                self.current_sheet_first_empty_column = 1
        except:
            print("Error getting first empty column")


    def __get_default_font_data(self):
        print("Getting default font data")
        try:
            return {
                "font": self.current_sheet.cell(row=1, column=1).font.name,
                "size": int(self.current_sheet.cell(row=1, column=1).font.size)
            }
        except:
            print("Error getting default font data")
            return {}
    
    def __get_cell_alignment(self, cell):
        print("Getting cell alignment")
        try:
            alignment = {}
            if cell.alignment:
                if cell.alignment.horizontal in ["center", "left", "right"]:
                    alignment["horizontal"] = cell.alignment.horizontal
                if cell.alignment.vertical in ["center", "bottom", "top"]:
                    alignment["vertical"] = cell.alignment.vertical
            return alignment
        except:
            print("Error getting cell alignment")
            return {}

    def __get_color_from_theme(self, color_data):
        print("Getting color from theme")
        try:
            color = theme_and_tint_to_rgb(self.workbook, color_data.theme, color_data.tint)
            return color
        except:
            print("Error getting color from theme")
            return {}

    def __get_color_data(self, color_data):
        print("Getting color data")
        try:
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
        except:
            print("Error getting color data")
            return None

    def __get_cell_font_data(self, cell):
        print("Getting cell font data")
        try:
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
        except:
            print("Error getting cell font data")
            return {}

    def __get_border_style(self, border_style):
        print("Getting border style")
        if border_style == "medium":
            return "thick"
        if border_style == "thick":
            return "extrathick"
        if border_style == "double":
            return "double"
        else:
          return "single"
    
    def __has_border(self, cell):
        print("Checking if cell has border")
        try:
            border = cell.border
            if border:
                if border.top:
                    if border.top.style:
                        return True
                if border.right:
                    if border.right.style:
                        return True
                if border.bottom:
                    if border.bottom.style:
                        return True
                if border.left:
                    if border.left.style:
                        return True
            return False
        except:
            print("Error checking if cell has border")
            return False
    
    def __get_cell_border_data(self, cell, is_merged_cell):
        print("Getting cell border data")
        try:
            cell_border_data = {}
            outline = {}

            border = cell.border

            if is_merged_cell:
                if border.top:
                    border_style = border.top.style
                    if border_style:
                        outline["style"] = self.__get_border_style(border_style)

                    border_color = border.top.color
                    if border_color:
                        outline["color"] = self.__get_color_data(border_color)
            else:
                border_top = []
                border_right = []
                border_bottom = []
                border_left = []

                for direction in ["top", "right", "bottom", "left"]:
                    if getattr(border, direction):
                        if getattr(border, direction).style:
                            locals()[f"border_{direction}"].append(self.__get_border_style(getattr(border, direction).style))
                        if getattr(border, direction).color:
                            locals()[f"border_{direction}"].append(self.__get_color_data(getattr(border, direction).color))

                if border_top == border_right == border_bottom == border_left and len(border_top) > 0:
                    border = border_top
                    if len(border) == 1:
                        cell_border_data["outline"] = {"style": border[0]}
                    if len(border) == 2:
                        cell_border_data["outline"] = {"style": border[0], "color": border[1]}
                else:
                    for direction in ["top", "right", "bottom", "left"]:
                        if locals()[f"border_{direction}"]:
                            if len(locals()[f"border_{direction}"]) == 1:
                                cell_border_data[direction] = {"style": locals()[f"border_{direction}"][0]}
                            if len(locals()[f"border_{direction}"]) == 2:
                                cell_border_data[direction] = {"style": locals()[f"border_{direction}"][0], "color": locals()[f"border_{direction}"][1]}

            if outline:
                cell_border_data["outline"] = outline
        
            return cell_border_data
        except:
            print("Error getting cell border data")
            return {}
    
    def __has_fill_color(self, cell):
        print("Checking if cell has fill color")
        try:
            if cell.fill:
                color = self.__get_color_data(cell.fill.start_color)
                if color and color != "#FFFFFF":
                    return True
            return False
        except:
            print("Error checking if cell has fill color")
            return False

    def __get_fill_color(self, cell):
        print("Getting cell fill color")
        try:
            if cell.fill:
                color = self.__get_color_data(cell.fill.start_color)
                if color and color != "#FFFFFF":
                    return color
            return None
        except:
            print("Error getting cell fill color")
            return None

    def __get_cell_data(self, cell):
        print("Getting cell data")
        print(cell.coordinate)
        try:
            cell_data = {
                "colnumber": cell.coordinate[0]
            }
            
            value = cell.value
            if value != None:
                if cell.is_date:
                    date = cell.value.strftime("%Y/%m/%d")
                    cell_data["value"] = date
                else:
                    cell_data["value"] = value

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
        except:
            print("Error getting cell data")
            return {}

    def __get_row_data(self, row):
        print("Getting row data")
        try:
            row_data = {
                "linenumber": row[0].row
            }

            columns = []
            for cell in row:
                if cell.column == self.current_sheet_first_empty_column:
                    break
                cell_data = self.__get_cell_data(cell)
                if cell_data:
                    columns.append(cell_data)
            
            if columns:
                row_data["columns"] = columns
            return row_data
        except:
            print("Error getting row data")
            return {}

    def __get_rows(self):
        print("Getting rows")
        try:
            rows = []
            for row in self.current_sheet.iter_rows():
                if row[0].row == self.current_sheet_first_empty_row:
                    break
                row_data = self.__get_row_data(row)
                if row_data:
                    rows.append(row_data)
            return rows
        except:
            print("Error getting rows")
            return []


    def __get_sheet_data(self, sheetnumber):
        print("Getting sheet data")
        try:
            font_data = self.__get_default_font_data()
            rows = self.__get_rows()

            sheet_data = {
                "sheetnumber": sheetnumber,
                "sheetname": self.current_sheet.title,
                "font": font_data,
                "lines": rows
            }

            return sheet_data
        except:
            print("Error getting sheet data")
            return {}

    def __open_workbook(self, excel_path):
        print("Opening workbook")
        try:
            workbook = openpyxl.load_workbook(excel_path)
            return workbook
        except:
            print("Error opening workbook")
            return None
    
    def parse_xlsx_to_json_file(self, excel_path):
        print("Parsing xlsx to json file")
        try:
            self.workbook = self.__open_workbook(excel_path)
            sheet_names = self.workbook.sheetnames

            sheet_data = []
            for i, sheetname in enumerate(sheet_names):
                self.current_sheet = self.workbook[sheetname]
                self.__set_first_empty_row()
                self.__set_first_empty_column()
                sheet_data.append(self.__get_sheet_data(i+1))

            return json.dumps({"sheets": sheet_data}, ensure_ascii=False)
        except Exception as e:
            print(e)
            return json.dumps({"error": str(e)}, ensure_ascii=False)
