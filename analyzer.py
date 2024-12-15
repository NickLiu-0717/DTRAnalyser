from odf.opendocument import load
from odf.table import Table, TableRow, TableCell
from odf.style import TableCellProperties
import pandas as pd
import matplotlib.colors as mcolors
import matplotlib.dates as mdates
import matplotlib.pyplot as plt
import matplotlib.font_manager as fm
import tkinter as tk
from tkinter import messagebox

class CustomAskWindow(tk.Toplevel):
    def __init__(self, parent, title, message, position="left"):
        super().__init__(parent)
        self.title(title)
        self.geometry("400x150")  # 設定自定義視窗的大小

        # 設定視窗位置
        screen_width = self.winfo_screenwidth()
        screen_height = self.winfo_screenheight()
        if position == "left":
            x = screen_width // 4 - 150  # 螢幕左側
        elif position == "right":
            x = screen_width * 3 // 4 - 150  # 螢幕右側
        else:
            x = (screen_width // 2) - 150  # 螢幕中間
        y = (screen_height // 2) - 75  # 垂直居中

        self.geometry(f"+{x}+{y}")  # 設置視窗位置
        self.resizable(False, False)  # 禁止調整大小

        # 問詢訊息
        tk.Label(self, text=message, font=("Arial", 12), wraplength=280).pack(pady=20)

        # Yes 和 No 按鈕
        self.result = None
        button_frame = tk.Frame(self)
        button_frame.pack()
        tk.Button(button_frame, text="Yes", width=10, command=self.on_yes).pack(side="left", padx=10)
        tk.Button(button_frame, text="No", width=10, command=self.on_no).pack(side="left", padx=10)

        # 設置視窗為模態（等待用戶操作）
        self.grab_set()
        self.wait_window(self)

    def on_yes(self):
        self.result = True  # 設定返回值為 True
        self.destroy()

    def on_no(self):
        self.result = False  # 設定返回值為 False
        self.destroy()

class Analyzer:
    def __init__(self, CACHE_FILE, ODS_FILE, root):
        if root != None:
            self.root = root
            self.root.title("Average Runs Calculator")
            
            self.root.pack_propagate(False)
            frame = tk.Frame(root)  # 新增框架，讓內容更緊湊
            frame.pack(padx=10, pady=10)  # 設置內外邊距，讓視窗看起來更美觀
            
            
            tk.Label(frame, text="Enter Start Date (YYYY-MM-DD):").grid(row=0, column=0, padx=10, pady=5)
            self.date1_entry = tk.Entry(frame, width=20)
            self.date1_entry.grid(row=0, column=1, padx=10, pady=5)
            
            tk.Label(frame, text="Enter End Date (YYYY-MM-DD):").grid(row=1, column=0, padx=10, pady=5)
            self.date2_entry = tk.Entry(frame, width=20)
            self.date2_entry.grid(row=1, column=1, padx=10, pady=5)
            
            self.calculate_button = tk.Button(frame, text="Calculate Average", command=self.calculate_average)
            self.calculate_button.grid(row=2, column=0, columnspan=2, pady=10)
            
            self.result_label = tk.Label(frame, text="Result will be displayed here.", fg="blue")
            self.result_label.grid(row=3, column=0, columnspan=2, pady=10)
        
        self.cach = CACHE_FILE
        self.ods = ODS_FILE
        self.expanded_table = []                                                 ## initial table for expanding by spanned row
        self.full_table = []                                                     ## initial table for expanding the whole table
        self.content_table = []                                                  ## initial tabel for only content we want
        self.indices = {"train": {}, "OP": {}, "DNF": {}, "FR": {}}              ## dictionary to store indices for expected cell contents
        self.odors_amounts = {"C": [], "A": [], "K": [], "H": [], "M": []}        ## dictionary to store amount of the odor usage
        self.odors_dates = {"C": [], "A": [], "K": [], "H": [], "M": []}         ## dictionary to store date of the odor usage
        self.warehouse = {}                                                      ## initial dictionary for warehouses
        self.flattened_data = []                                                 ## initial blank list for plot the warehouse graph
        self.duty_dates = []                                                     ## initial blank list for duty dates  
        self.duty_times = []                                                     ## initial blank list for duty times
        self.__load_data()                                                       ## load the data and get the expanded table
        self.__expand_column_repeated()                                          ## get the full table 
        self.__get_content_index()
        self.__get_warehouse_statistic()
        self.__get_train_odor_trend()   
    
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
                    nums = str(row[4]).split("/")
                    num = int(nums[0])
                    self.warehouse[str(row[1])][str(row[2])][str(row[3])] = num
            
            else:
                if str(row[2]) not in self.warehouse[str(row[1])]:
                    if str(row[4]) != "0 / 0 = 0":
                        self.warehouse[str(row[1])][str(row[2])] = {}
                        nums = str(row[4]).split("/")
                        num = int(nums[0])
                        self.warehouse[str(row[1])][str(row[2])][str(row[3])] = num
                else:
                    if str(row[3]) not in self.warehouse[str(row[1])][str(row[2])]:
                        if row[4] != "0 / 0 = 0":
                            nums = str(row[4]).split("/")
                            num = int(nums[0])
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
        
        unique_colors = list(mcolors.TABLEAU_COLORS.values()) + list(mcolors.CSS4_COLORS.values())
        unique_colors = unique_colors[:len(summary.columns)]  # 確保顏色數量足夠
        # Plot
        # summary.plot(kind="bar", stacked=True, figsize=(10, 6))
        ax = summary.plot(kind="bar", stacked=True, figsize=(10, 6), color=unique_colors)
        plt.title("Total Imports and Exports by Category")
        plt.ylabel("Value")
        plt.xlabel("Warehouse")
        plt.xticks(fontproperties=font_prop, rotation=0)
        plt.legend(title="Type", prop=font_prop)
        plt.tight_layout()
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
            if index1 > index2:
                return "ERROR: The Start Date Should be Before the End Date."
        elif date1 not in str_duty_date and date2 not in str_duty_date:
            return "ERROR: Both Dates are not Valid."
        elif date1 not in str_duty_date:
            return "ERROR: Start Date is not Valid."
        elif date2 not in str_duty_date:
            return "ERROR: End Date is not Valid."
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
        if count_days > 0:
            average_runs = round(sum_runs / count_days, 2)
            return f"The average runs between {date1} and {date2} is {average_runs}"
        else:
            return "No valid data found for the given dates."

    def calculate_average(self):
        date1 = self.date1_entry.get()
        date2 = self.date2_entry.get()
        
        # 驗證輸入日期並計算結果
        if not date1 or not date2:
            messagebox.showerror("Input Error", "Please enter both dates.")
            return
        
        result = self.get_average_runs_for_dates(date1, date2)
        self.result_label.config(text=result)
       
    def __get_train_odor_trend(self):
        for row_index, col_list in self.indices["train"].items():
            for col_index in col_list:
                cell = self.content_table[row_index + 1][col_index]
                odor_list = str(cell).split()
                if odor_list:
                    if odor_list[0] == "H":
                        split_amount = odor_list[-1].split("*")
                        odor_list[-1] = str(int(split_amount[0]) * int(split_amount[1]))
                if len(odor_list) > 2:
                    if odor_list[0] in self.odors_amounts:
                        self.odors_amounts[odor_list[0]].append(int(odor_list[-1]))
                        self.odors_dates[odor_list[0]].append(str(self.duty_dates[col_index - 6]))
            
    def plot_odor_trends(self):
        num_keys = len(self.odors_amounts)
        # 創建多個子圖：1列多行
        fig, axes = plt.subplots(nrows=num_keys, ncols=1, figsize=(6, 2 * num_keys), sharex=True)

        # 如果只有一個 key，axes 不是列表，需要轉換
        if num_keys == 1:
            axes = [axes]

        # 定義顏色列表，使用 TABLEAU 顏色或 CSS4 顏色
        color_list = list(mcolors.TABLEAU_COLORS.values())

        for ax, (key, values), color in zip(axes, self.odors_amounts.items(), color_list):
            dates = self.odors_dates[key]

            # 檢查兩個列表是否對應
            if len(values) != len(dates):
                print(f"Key '{key}' 的數據和日期長度不匹配，跳過繪圖。")
                continue

            # 創建 DataFrame 並排序
            df = pd.DataFrame({'Date': pd.to_datetime(dates), 'Value': values})
            df = df.sort_values(by='Date')  # 按日期排序

            # 繪製子圖，使用不同的顏色
            ax.plot(df['Date'], df['Value'], marker='o', label=key, color=color)
            ax.set_title(f"Trend for {key}")
            ax.set_ylabel("Value")
            ax.grid(True)
            ax.legend()

            # 設置 X 軸的日期格式
            ax.xaxis.set_major_locator(mdates.DayLocator(interval=15))  # 自動調整刻度
            ax.xaxis.set_major_formatter(mdates.DateFormatter("%Y-%m-%d"))  # 日期格式

        # 統一 X 軸格式
        plt.xlabel("Date")
        plt.xticks(rotation=45)
        plt.tight_layout()  # 自動調整佈局
        plt.show()
 