import openpyxl
import json
from libs.color_helper import theme_and_tint_to_rgb
import zipfile
import xml.etree.ElementTree as ET
import sys

class ExcelParser:
    workbook = None
    custom_index  = None
    current_sheet = None
    current_range = None
    empty_rows = 0
    empty_columns = 0

    def __get_merged_ranges(self):
        return self.current_sheet.merged_cells.ranges if self.current_sheet.merged_cells else []

    def __get_first_cells_of_merged_ranges(self):
        return [merged_range.start_cell.coordinate for merged_range in self.__get_merged_ranges()]

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

    def __get_default_font_data(self):
        try:
            return {
                "font": self.current_sheet.cell(row=1, column=1).font.name,
                "size": int(self.current_sheet.cell(row=1, column=1).font.size)
            }
        except:
            print("Error getting default font data")
            return {}
    
    def __get_cell_alignment(self, cell):
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
        try:
            color = theme_and_tint_to_rgb(self.workbook, color_data.theme, color_data.tint)
            return color
        except:
            print("Error getting color from theme")
            return {}

    def __get_color_data(self, color_data):
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
                    if self.custom_index:
                        Colors = self.custom_index
                    else:
                        Colors = openpyxl.styles.colors.COLOR_INDEX
                    color = Colors[color_data.indexed][2:]
            if color_data.type == "theme":
                color = self.__get_color_from_theme(color_data)
            return f"#{color}"
        except:
            print("Error getting color data (" + str(color_data) + ")") 
            return None

    def __get_cell_font_data(self, cell):
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
        if border_style == "medium":
            return "thick"
        if border_style == "thick":
            return "extrathick"
        if border_style == "double":
            return "double"
        else:
          return "single"

    def __set_border(self, cell, border, direction, partner):
        if getattr(border, direction) and getattr(border, direction).style and getattr(border, direction).color:
            locals()[f"border_{direction}"].append(self.__get_border_style(getattr(border, direction).style))
            locals()[f"border_{direction}"].append(self.__get_color_data(getattr(border, direction).color))
        else:
            if cell.row == 1 and direction == "top":
                return False
            if cell.column == 1 and direction == "left":
                return False
            if direction == "top":
                neighbor = self.current_sheet.cell(row=cell.row-1, column=cell.column)
            if direction == "right":
                neighbor = self.current_sheet.cell(row=cell.row, column=cell.column+1)
            if direction == "bottom":
                neighbor = self.current_sheet.cell(row=cell.row+1, column=cell.column)
            if direction == "left":
                neighbor = self.current_sheet.cell(row=cell.row, column=cell.column-1)
            if getattr(neighbor.border, partner) and getattr(neighbor.border, partner).style and getattr(neighbor.border, partner).color:
                locals()[f"border_{direction}"].append(self.__get_border_style(getattr(neighbor.border, partner).style))
                locals()[f"border_{direction}"].append(self.__get_color_data(getattr(neighbor.border, partner).color))
        return locals()[f"border_{direction}"]
    
    def __get_cell_border_data(self, cell, is_merged_cell):
        try:
            cell_border_data, outline = {}, {}

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
                [border_top, border_right, border_bottom, border_left]=[self.__set_border(cell, border, direction, partner) for direction, partner in {"top": "bottom", "right": "left", "bottom": "top", "left": "right"}.items() 

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
            print("Error getting cell border data (" + cell.coordinate + ")")
            return {}

    def __get_fill_color(self, cell):
        try:
            if cell.fill:
                if cell.fill.start_color:
                    color = self.__get_color_data(cell.fill.start_color)
                    if color and color != "#FFFFFF":
                        return color
            return None
        except:
            print("Error getting cell fill color (" + str(cell.coordinate) + ")")
            return None

    def __get_cell_data(self, cell):
        try:
            cell_data = {
                "colnumber": cell.coordinate[0]
            }
            
            value = cell.value
            if value != None:
                cell_data["value"] = cell.value.strftime("%Y/%m/%d") if cell.is_date else value

            is_merged_cell = self.__is_merged_cell(cell, self.__get_merged_ranges())
            if is_merged_cell:
                if cell.coordinate not in self.__get_first_cells_of_merged_ranges():
                    return {}
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
            
            if not cell_data.get("value") or not cell_data.get("border") or not cell_data.get("fill"):
                return {}

            return cell_data
        except:
            print("Error getting cell data (" + cell.coordinate + ")")
            print(sys.exc_info()[0])
            return {}

    def __get_cell_data_wrapper(self, cell):
        cell_data = self.__get_cell_data(cell)
        self.empty_columns = self.empty_columns + 1 if cell_data == {} else 0
        return cell_data

    def __get_row_data(self, row):
        self.empty_columns = 0
        try:
            row_data = {
                "linenumber": row[0].row
            }

            columns = [self.__get_cell_data_wrapper(cell) for cell in row if self.empty_columns < 50 and cell.column < self.current_sheet.max_column+2]
            if columns:
                row_data["columns"] = columns

            return row_data
        except:
            print("Error getting row data")
            return {}

    def __get_row_data_wrapper(self, row):
        row_data = self.__get_row_data(row)
        self.empty_rows = self.empty_rows + 1 if row_data.get('columns') is None or row_data == {} else 0
        return row_data


    def __get_rows(self):
        self.empty_rows = 0
        try:
            rows = [self.__get_row_data_wrapper(row) for row in self.current_sheet.iter_rows() if self.empty_rows < 50 and row[0].row < self.current_sheet.max_row+2]
            return rows
        except:
            print("Error getting rows")
            return []


    def __get_sheet_data(self, sheetnumber):
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

    def __set_custom_index(self, color):
        self.custom_index = [rgb.attrib['rgb'] for rgb in color]

    def __check_for_custom_index(self, filepath):
        try:
            with zipfile.ZipFile(filepath) as zgood:
                styles_xml = zgood.read('xl/styles.xml')
                root = ET.fromstring(styles_xml)
                [[self.__set_custom_index(color) for color in child if 'indexedColors' in color.tag] for child in root if 'colors' in child.tag]
 
        except:
            print("Error checking for custom index")
            return None
        

    def __open_workbook(self, excel_path):
        try:
            self.workbook = openpyxl.load_workbook(excel_path, data_only=True)
        except:
            print("Error opening workbook")
            return None

    def __get_sheet(self, i, sheetname):
        self.current_sheet = self.workbook[sheetname]
        return self.__get_sheet_data(i+1)
    
    def parse_xlsx_to_json_file(self, excel_path):
        try:
            self.__open_workbook(excel_path)
            self.__check_for_custom_index(excel_path)
            sheet_names = self.workbook.sheetnames

            sheet_data = [self.__get_sheet(i, sheetname) for i, sheetname in enumerate(sheet_names)]

            return json.dumps({"sheets": sheet_data}, ensure_ascii=False)
        except Exception as e:
            print(e)
            return json.dumps({"error": str(e)}, ensure_ascii=False)
