import os
import openpyxl
from openpyxl import styles # Could not use openpyxl to load data as the budget reports is of .xls format
from openpyxl.styles import Font, PatternFill, Border, Side
from openpyxl.utils import get_column_letter
import xlrd
import pandas as pd
from pathlib import Path
from operator import itemgetter

class AutoBudget():

    def __init__(self, input_dir_path) -> None:
        # Init
        self.workbook = openpyxl.Workbook() # Create workbook
        self.cost_center_set = set()
        self.budget_dict = self.load_budgets(input_dir_path)

        # Styles
        self.font_bold = Font(name='Arial', bold=True, size=13)
        self.font_small_bold = Font(name='Arial', bold=True, size=11)
        self.font_standard = Font(name='Arial')
        self.color_light_blue = PatternFill(fgColor='caddeb', fill_type='solid')
        self.color_dark_blue = PatternFill(fgColor='8db0db', fill_type='solid')
        self.color_light_red = PatternFill(fgColor='d8bdc2', fill_type='solid')
        self.color_dark_red = PatternFill(fgColor='d99ea9', fill_type='solid')


        self.border_thin = Border(
            left=Side(border_style='thin', color='CDCDCD'), 
            right=Side(border_style='thin', color='CDCDCD'),
            top=Side(border_style='thin', color='CDCDCD'),
            bottom=Side(border_style='thin', color='CDCDCD')
            )
        self.border_thick_right = Border(
            left=Side(border_style='thin', color='CDCDCD'), 
            right=Side(border_style='thick', color='000000'),
            top=Side(border_style='thin', color='CDCDCD'),
            bottom=Side(border_style='thin', color='CDCDCD')
            )
        self.border_thick_bottom = Border(
            left=Side(border_style='thin', color='CDCDCD'), 
            right=Side(border_style='thin', color='CDCDCD'),
            top=Side(border_style='thin', color='CDCDCD'),
            bottom=Side(border_style='thick', color='000000')
            )
        self.border_thick_double = Border(
            left=Side(border_style='thin', color='CDCDCD'), 
            right=Side(border_style='thick', color='000000'),
            top=Side(border_style='thin', color='CDCDCD'),
            bottom=Side(border_style='thick', color='000000')
            )

        self.offset = 2 # Offset from 0 where table begins


    def load_budgets(self, input_dir_path) -> dict:
        budget_dict = {}
        doublet_check_list = []
        for file in os.listdir(input_dir_path):
            if file.endswith(".XLS"):
                file_path = os.path.join(input_dir_path, file)
                wb = xlrd.open_workbook(Path(file_path))
                cost_report = wb["Cost center report"]

                # Find date, cost_center and table index
                for index, row in enumerate(cost_report.get_rows()):
                    info_cost_elemnt = row[5].value
                    if info_cost_elemnt == "Fiscal period / year (Interval, Req.)":
                        date = row[6].value
                    elif info_cost_elemnt == "Cost Center Node":
                        cost_center = row[6].value[-4:]
                    elif info_cost_elemnt == "Table":
                        start_of_table = index + 1
                    elif info_cost_elemnt == "HSQVBI_CCTR_GR":
                        end_of_table = index - 1

                # Make table array
                table = [cost_report.row_values(row) for row in range(start_of_table, end_of_table)]

                # Make dataframe
                columns = table[0]
                columns[6] = "Cost Type"
                table_df = pd.DataFrame(table[1:], columns=table[0])
                table_df = table_df.set_index("Cost Type")

                # Remove unwanted columns and rows
                table_df = table_df.drop(columns=["Cost Element", ""])
                table_df = table_df.drop("")

                if not date in budget_dict:
                    budget_dict[date] = []

                check = cost_center + date
                if check in doublet_check_list:
                    raise Exception(f"cost center {cost_center} has a dublicate with the date {date}, please remove it!")
                budget_dict[date].append({cost_center : table_df, 'id' : cost_center})
                doublet_check_list.append(check)
                self.cost_center_set.add(cost_center)

        return budget_dict

    # Sum different costs based on actual cost
    # pos: cost = 0 budget = 1 
    def sum_month(self, cost_center_list, pos) -> dict:
        month_dict = {}
        for dict in cost_center_list:
            df = list(dict.values())[0]
            for index, row in df.iterrows():
                cost = row[pos]
                if not cost:
                    cost = 0
                if not index in month_dict:
                    month_dict[index] = float(cost)
                else:
                    month_dict[index] += float(cost)
        return month_dict
        

    def make_compilation(self):
        compilation_sheet = self.workbook.active
        compilation_sheet.title = "Sammanställning Kostnadsslag"
        same_every_col = len(self.cost_center_set)+2

        # Add standard headers to worksheet
        month_header = ["Jan", "Feb", "Mar", "Apr", "May", "Jun", "Jul", "Aug", "Sep", "Oct", "Nov", "Dec"]
        for i in range(len(month_header)):
            compilation_sheet.cell(self.offset,i*same_every_col+self.offset+1, month_header[i]).font = self.font_bold
            compilation_sheet.cell(self.offset,i*same_every_col+self.offset+2, "Budget").font = self.font_small_bold
        compilation_sheet.cell(self.offset,self.offset, compilation_sheet.title).font = self.font_bold

        # Add costs and cost types to worksheet
        existing_cost_types = {}
        for date, cost_center_list in self.budget_dict.items():
            i_col = (int(date[:3])-1)*same_every_col + self.offset+1

            # Actual & cost types
            month_cost_dict_actual = self.sum_month(cost_center_list, 0)
            for cost_type, cost in month_cost_dict_actual.items():
                if cost_type in existing_cost_types.keys():
                        compilation_sheet.cell(existing_cost_types.get(cost_type), i_col, cost).font = self.font_standard
                        compilation_sheet.cell(existing_cost_types.get(cost_type), i_col, cost).style = 'Comma [0]'
                else:
                    i_row = compilation_sheet.max_row + 1
                    compilation_sheet.cell(i_row, self.offset, cost_type).font = self.font_standard # Add cost type
                    existing_cost_types[cost_type] = i_row
                    compilation_sheet.cell(i_row, i_col, cost).font = self.font_standard
                    compilation_sheet.cell(i_row, i_col, cost).style = 'Comma [0]'

            # Planned
            month_cost_dict_planned = self.sum_month(cost_center_list, 1)
            for cost_type, cost in month_cost_dict_planned.items():
                compilation_sheet.cell(existing_cost_types.get(cost_type), i_col+1, cost).font = self.font_standard
                compilation_sheet.cell(existing_cost_types.get(cost_type), i_col+1, cost).style = 'Comma [0]'

            # Add cost for individual cost centers
            cost_center_list = sorted(cost_center_list, key=itemgetter('id'))
            for i, cost_center_dict in enumerate(cost_center_list):
                offset_individual = i + 2
                #for cost_center, df in cost_center_dict.items():
                compilation_sheet.cell(self.offset, i_col + offset_individual, cost_center_dict['id']).font = self.font_small_bold
                for cost_type, row in cost_center_dict[cost_center_dict['id']].iterrows():
                    actual_cost = row[0]
                    if actual_cost:
                        compilation_sheet.cell(existing_cost_types.get(cost_type), i_col + offset_individual, actual_cost).font = self.font_standard
                        compilation_sheet.cell(existing_cost_types.get(cost_type), i_col + offset_individual, actual_cost).style = 'Comma [0]'

        self.make_sum_rows(compilation_sheet, same_every_col)
        self.style_sheet(compilation_sheet, same_every_col)

    def make_sum_rows(self, sheet, same_every_col):

        # Sum all columns
        row_total = sheet.max_row+2
        sheet.cell(row_total, self.offset, "Totalt").font = self.font_small_bold
        for col in range(self.offset+1,sheet.max_column+1):
            column_letter = get_column_letter(col)
            sheet.cell(row_total, col, f"=SUM({column_letter}{2}:{column_letter}{row_total-2})").font = self.font_small_bold
            sheet.cell(row_total, col, f"=SUM({column_letter}{2}:{column_letter}{row_total-2})").style = 'Comma [0]'

        # Accumulation of sums
        row = sheet.max_row+1
        sheet.cell(row, self.offset, "Totalt (ACC)").font = self.font_small_bold
        for col in range(self.offset+1, sheet.max_column+1, same_every_col):
            column_letter = get_column_letter(col)
            if col == self.offset+1:
                sheet.cell(row, col, f"=SUM({column_letter}{row-1}+0)").font = self.font_small_bold
                sheet.cell(row, col, f"=SUM({column_letter}{row-1}+0)").style = 'Comma [0]'
            else:
                sheet.cell(row, col, f"=SUM({column_letter}{row-1}+{get_column_letter(col-same_every_col)}{row})").font = self.font_small_bold
                sheet.cell(row, col, f"=SUM({column_letter}{row-1}+{get_column_letter(col-same_every_col)}{row})").style = 'Comma [0]'
        
        # Move budget sums
        row = sheet.max_row+1
        sheet.cell(row, self.offset, "Totalt Budget").font = self.font_small_bold
        for col in range(self.offset+2, sheet.max_column+1, same_every_col):
            budget_sum = sheet.cell(row_total, col).value
            sheet.cell(row_total, col, "-").font = self.font_standard
            sheet.cell(row_total, col, "-").style = 'Comma [0]'
            sheet.cell(row, col-1, budget_sum).font = self.font_small_bold
            sheet.cell(row, col-1, budget_sum).style = 'Comma [0]'

    def autosize_column(self, ws, columnrange, length = 0):
        for column in columnrange:
            column_cells = [c for c in ws.columns][column-1]
            if not length:
                length = max(len(str(cell.value))*0.87 for cell in column_cells)
            ws.column_dimensions[column_cells[0].column_letter].width = length

    #Set thich border around given range
    def set_thick_border(self, sheet, startRow, startCol, endRow, endCol):
        max_y = endRow - startRow  # index of the last row
        for pos_y, r in enumerate(range(startRow, endRow +1)):
            max_x = endCol - startCol  # index of the last cell
            for pos_x, c in enumerate(range(startCol, endCol +1)):
                cell = sheet.cell(r,c)
                BORDER = Border(
                    left=cell.border.left,
                    right=cell.border.right,
                    top=cell.border.top,
                    bottom=cell.border.bottom
                    )
                if pos_x == 0:
                    BORDER.left = Side(border_style='thick', color='000000')
                if pos_x == max_x:
                    BORDER.right = Side(border_style='thick', color='000000')
                if pos_y == 0:
                    BORDER.top = Side(border_style='thick', color='000000')
                if pos_y == max_y:
                    BORDER.bottom = Side(border_style='thick', color='000000')
                # set new border only if it's one of the edge cells
                if pos_x == 0 or pos_x == max_x or pos_y == 0 or pos_y == max_y:
                    cell.border = BORDER

    def style_sheet(self, sheet, same_every_col):
        # Set column colors
        color = 0
        for i, col in enumerate(sheet.columns):
            column_index = i+1
            if column_index == self.offset:
                # Decide this color later
                pass
            elif column_index > self.offset:
                if (column_index - self.offset -1)%same_every_col == 0:
                    color += 1
                    for cell in col:
                        if cell.row >= self.offset:
                            if color%2 == 0:
                                cell.fill = self.color_dark_red
                            else:
                                cell.fill = self.color_dark_blue
                else:
                    for cell in col:
                        if cell.row >= self.offset:
                            if color%2 == 0:
                                cell.fill = self.color_light_red
                            else:
                                cell.fill = self.color_light_blue

        # Set column borders
        for col in range(self.offset, sheet.max_column + 1):
            for row in range(self.offset, sheet.max_row + 1):
                sheet.cell(row, col).border = self.border_thin
                if col ==  self.offset and row == self.offset:
                    sheet.cell(row, col).border = self.border_thick_double
                elif col == self.offset:
                    sheet.cell(row, col).border = self.border_thick_right
                elif row == self.offset:
                    sheet.cell(row, col).border = self.border_thick_bottom
        self.set_thick_border(sheet, self.offset, self.offset, sheet.max_row, sheet.max_column)

        # Hide columns
        sheet.sheet_properties.outlinePr.summaryRight = False
        for i in range(same_every_col-1): # Grouping several at once does not work, but one at a time works
            for col in range(self.offset+2+i, sheet.max_column, same_every_col):
                sheet.column_dimensions.group(get_column_letter(col), get_column_letter(col), hidden=True) 

        # Size columns
        self.autosize_column(sheet, [self.offset])
        self.autosize_column(sheet, range(self.offset+1,sheet.max_column), 10)


if __name__ == "__main__":

    print("Hello, I am your budget automator")
    budget = AutoBudget("./data/")
    budget.make_compilation()
    budget.workbook.save("Sammanställning.xlsx")