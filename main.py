# from ExpandTable import *
# from CellProcess import *
# from GrabContent import *
from odf.style import Style
import pandas as pd
import matplotlib.pyplot as plt
import matplotlib.font_manager as fm
from analyzer import Analyzer, CustomAskWindow
import tkinter as tk
from tkinter import messagebox

def column_letter_to_index(column_letter):
    """
    將英文字母列標記 (A, B, ..., Z, AA, AB, ...) 轉換為列索引 (1, 2, ..., 26, 27, ...)
    """
    column_letter = column_letter.upper()  # 確保是大寫
    index = 0
    for char in column_letter:
        index = index * 26 + (ord(char) - ord('A') + 1)
    return index - 1

def center_window_on_primary(window, width, height, offset="left"):
    """
    讓視窗在主螢幕的左側或右側顯示，避免跨螢幕居中。
    
    :param window: tkinter 視窗物件
    :param width: 視窗寬度
    :param height: 視窗高度
    :param offset: 偏移方向 ("left" 或 "right")
    """
    screen_width = window.winfo_screenwidth()  # 主螢幕寬度
    screen_height = window.winfo_screenheight()  # 主螢幕高度
    
    # 根據 offset 決定視窗的 X 座標
    if offset == "left":
        x = screen_width // 4 - width // 2  # 螢幕左側 1/4 處
    elif offset == "right":
        x = screen_width * 3 // 4 - width // 2  # 螢幕右側 3/4 處
    else:
        x = (screen_width // 2) - (width // 2)  # 預設居中
    
    y = (screen_height // 2) - (height // 2)  # 垂直居中
    
    window.geometry(f"{width}x{height}+{x}+{y}")  # 設定視窗大小和位置
    window.update()

def main():
    CACHE_FILE = "cached_ods.pkl"
    # ODS_FILE = "2023-3months-DTR.ods"
    ODS_FILE = "2023-oneandhalfyear-DTR.ods"
    root = tk.Tk()
    root.withdraw()
    # center_window_on_primary(root, 600, 200, offset="left")
    # response = messagebox.askyesno("Start Application", "Do you want to calculate average runs?")
    custom_ask = CustomAskWindow(root, "Start Application", "Do you want to calculate average runs?", position="left")
    if custom_ask.result:  # 如果選擇「是」
        root.deiconify()
        center_window_on_primary(root, 600, 200, offset="left")
        dtr_analyzer = Analyzer(CACHE_FILE, ODS_FILE, root)
        root.update()  # 更新視窗內容
        root.geometry("600x200")  # 自動調整視窗大小
        root.mainloop()
        # dtr_analyzer.plot_warehouse()
        # dtr_analyzer.plot_odor_trends()
    else:  # 如果選擇「否」
        root.destroy()
        dtr_analyzer = Analyzer(CACHE_FILE, ODS_FILE, root=None)
        # dtr_analyzer.plot_warehouse()
        # dtr_analyzer.plot_odor_trends()

    
    # dtr_analyzer.get_average_runs_for_dates("2023/10/11", "2024/1/10")  
    # dtr_analyzer.plot_warehouse()
    # dtr_analyzer.plot_odor_trends()
    


main()

