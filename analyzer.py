from odf.opendocument import load
from odf.table import Table, TableRow, TableCell
from odf.style import TableCellProperties
import cloudpickle, re, sys
from useful_func import *
import pandas as pd
import matplotlib.pyplot as plt
import matplotlib.font_manager as fm

sys.setrecursionlimit(4000)

class Analyzer:
    def __init__(self, CACHE_FILE, ODS_FILE):
        self.cach = CACHE_FILE
        self.ods = ODS_FILE
        self.expanded_table = []                                                 ## initial table for expanding by spanned row
        self.full_table = []                                                     ## initial table for expanding the whole table
        self.content_table = []                                                  ## initial tabel for only content we want
        self.indices = {"train": {}, "OP": {}, "DNF": {}, "FR": {}}              ## indices for expected cell contents
        self.warehouse = {}                                                      ## initial dictionary for warehouses
        self.flattened_data = []                                                 ## initial blank list for plot the warehouse graph
        self.duty_dates = []                                                     ## initial blank list for duty dates  
        self.duty_times = []                                                     ## initial blank list for duty times
        self.__load_data()                                                       ## load the data and get the expanded table
        self.__expand_column_repeated()                                          ## get the full table 
    
    def run_all_step(self):
        self.__get_content_index()
        self.__get_warehouse_statistic()
        self.plot_warehouse()
    
    def get_cell_from_index(self, row, col):
        print(f"Index:{(row, col)}, Cell Content:{str(self.full_table[row][col])}")
    
    def __load_data(self):
        print("Loading data from ODS file...")
        doc = load(self.ods)
        self.expanded_table = self.__expand_row_merged_cells(doc)
    
    def __find_gap(self, lst):
        lst.sort()
        for i in range(1, len(lst)):
            if lst[i] - lst[i - 1] > 1:
                return i
        return None
    
    def __expand_row_merged_cells(self, doc):
        sheet = doc.spreadsheet.getElementsByType(Table)[0]
        expanded_table = []
        track_index = {}
        for row_index, row in enumerate(sheet.getElementsByType(TableRow)):
            while len(expanded_table) <= row_index:
                expanded_table.append([])
            col_index = 0
            for cell in row.getElementsByType(TableCell):
                cell_value = ""
                for child in cell.childNodes:
                    if child.tagName == "text:p":
                        cell_value = "".join(text_node.data for text_node in child.childNodes if text_node.nodeType == 3)
                if cell_value is None:
                    cell_value = ""
                row_span = cell.getAttribute("numberrowsspanned")
                row_span = int(row_span) if row_span is not None else 1       
                for i in range(row_span):
                    if row_index + i not in track_index:
                        track_index[row_index + i] = []
                    while len(expanded_table) < row_index + row_span:  # 確保目標行存在
                        expanded_table.append([])
                    if col_index < len(expanded_table[row_index + i]):
                        gap_index = self.__find_gap(track_index[row_index + i])
                        if gap_index:
                            col_index = gap_index
                        else: 
                            if expanded_table[row_index + i][col_index] != "":
                                col_index = len(expanded_table[row_index + i])
                    while len(expanded_table[row_index + i]) <= col_index:  # 確保列存在
                        expanded_table[row_index + i].append("")               
                    expanded_table[row_index + i][col_index] = cell
                    track_index[row_index + i].append(col_index)
                col_index += 1
        return expanded_table
    
    def __expand_column_repeated(self):   
        self.full_table = [[] for _ in range(len(self.expanded_table))]
        for row in range(len(self.expanded_table)):
            for col, cell in enumerate(self.expanded_table[row]):
                try:
                    col_span = cell.getAttribute("numbercolumnsspanned")
                    col_span = int(col_span) if col_span is not None else 1
                    
                    col_repeat = cell.getAttribute("numbercolumnsrepeated")
                    col_repeat = int(col_repeat) if col_repeat is not None else 1
                except AttributeError as e:
                    
                    print(f"Error {e} at index: ({(row, col)})")
                if  col_span > 1: 
                    for _ in range(col_span):
                        self.full_table[row].append(cell)
                    continue
                if col_repeat < 200:
                    for _ in range(col_repeat):
                        self.full_table[row].append(cell)
        self.content_table = self.full_table[84:654]
        self.duty_dates = self.full_table[0][6:]                                                        
        self.duty_times = self.full_table[2][6:]

    def __extract_text_from_element(self, element, layer=0):
        """
        遞迴提取節點內的所有文本內容。   layer要修改 要提取出日期
        """
        text = ""
        for child in element.childNodes:
            layer += 1
            if hasattr(child, "data"):  # 如果是文本節點
                if layer != 2:
                    text += child.data
            elif hasattr(child, "childNodes"):  # 如果是元素節點，遞迴處理
                text += self.__extract_text_from_element(child, layer)
        return text
    
    def __get_cell_content_and_annotation(self, cell):
        cell_content = ""
        annotation_text = ""
        for child in cell.childNodes:
            if child.tagName == "text:p":  # 確保是文本段落
                cell_content += "".join(text_node.data for text_node in child.childNodes if text_node.nodeType == 3)
            if child.tagName == "office:annotation":
                annotation_text = self.__extract_text_from_element(child)
        return cell_content, annotation_text
    
    def __filter_func(self, cell):
        c, a = self.__get_cell_content_and_annotation(cell)
        if c == "*":
            return "train"
        if c == "●":
            return "OP"
        if c == "DNF":
            if not a:
                return "DNF"
        if c == "FR":
            return "FR"
    
    def __get_content_index(self):      ## method that get self.indices
        for r, row in enumerate(self.content_table):
            for i, cell in enumerate(row):
                category = self.__filter_func(cell)
                if category:
                    if r not in self.indices[category]:
                        self.indices[category][r] = []
                    self.indices[category][r].append(i)
    
    def __get_warehouse_statistic(self): ##0 / 0 = 0  method that get self.warehouse
        for row in self.content_table:
            if str(row[1]) not in self.warehouse:
                if str(row[4]) != "0 / 0 = 0":
                    self.warehouse[str(row[1])] = {}
                    self.warehouse[str(row[1])][str(row[2])] = {}
                    num = int(str(row[4])[0])
                    self.warehouse[str(row[1])][str(row[2])][str(row[3])] = num
            
            else:
                if str(row[2]) not in self.warehouse[str(row[1])]:
                    if str(row[4]) != "0 / 0 = 0":
                        self.warehouse[str(row[1])][str(row[2])] = {}
                        num = int(str(row[4])[0])
                        self.warehouse[str(row[1])][str(row[2])][str(row[3])] = num
                else:
                    if str(row[3]) not in self.warehouse[str(row[1])][str(row[2])]:
                        if row[4] != "0 / 0 = 0":
                            num = int(str(row[4])[0])
                            self.warehouse[str(row[1])][str(row[2])][str(row[3])] = num
    
    def __flatten_dict(self, d, keys=[]):   
        for k, v in d.items():
            if isinstance(v, dict):
                self.__flatten_dict(v, keys + [k])
            else:
                self.flattened_data.append(keys + [k, v])
    
    def plot_warehouse(self):
        font_path = "msjhl.ttc"
        font_prop = fm.FontProperties(fname=font_path)
        self.__flatten_dict(self.warehouse)
        columns = ["Warehouse", "Subcategory", "Detail", "Value"]
        df = pd.DataFrame(self.flattened_data, columns=columns)
        summary = df.groupby(["Warehouse", "Subcategory"])["Value"].sum().unstack()
        summary["Total"] = summary.sum(axis=1)  # Add a 'Total' column for sorting
        summary = summary.sort_values("Total", ascending=False).drop(columns=["Total"])  # Sort and drop 'Total'
        # Plot
        summary.plot(kind="bar", stacked=True, figsize=(10, 6))
        plt.title("Total Imports and Exports by Category")
        plt.ylabel("Value")
        plt.xlabel("Warehouse")
        plt.xticks(fontproperties=font_prop)
        plt.legend(title="Type", prop=font_prop)
        plt.show()
        
    def __get_runs(self):
        runs_dict = {}
        for i, run in enumerate(self.duty_times):
            if str(run):
                if len(str(run)) > 1:
                    if str(run)[0].isdigit() and str(run)[1] == "(":
                        runs_dict[i] = 2 * int(str(run)[0])
                else:
                    if str(run)[0].isdigit():
                        runs_dict[i] = int(str(run)[0])
        return runs_dict

    def get_average_runs_for_dates(self, date1, date2):
        str_duty_date = [str(date) for date in self.duty_dates]
        if date1 in str_duty_date and date2 in str_duty_date:
            index1 = str_duty_date.index(date1)
            index2 = str_duty_date.index(date2)
        else:
            print("Input dates are not in the list")
        runs_dict = self.__get_runs()
        key = index1
        count_days = 0
        sum_runs = 0
        while key >= index1 and key <= index2:
            if key in runs_dict:
                sum_runs += runs_dict[key]
                count_days += 1
                key += 1
            else:
                key += 1
        print(f"The average runs between {date1} and {date2} is {round(sum_runs / count_days, 2)}")
 