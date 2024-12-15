# from ExpandTable import *
# from CellProcess import *
# from GrabContent import *
from odf.style import Style
import pandas as pd
import matplotlib.pyplot as plt
import matplotlib.font_manager as fm
from analyzer import Analyzer

def column_letter_to_index(column_letter):
    """
    將英文字母列標記 (A, B, ..., Z, AA, AB, ...) 轉換為列索引 (1, 2, ..., 26, 27, ...)
    """
    column_letter = column_letter.upper()  # 確保是大寫
    index = 0
    for char in column_letter:
        index = index * 26 + (ord(char) - ord('A') + 1)
    return index - 1

def main():
    CACHE_FILE = "cached_ods.pkl"
    # ODS_FILE = "2023-3months-DTR.ods"
    ODS_FILE = "2023-oneandhalfyear-DTR.ods"
    # expand_row_table = load_data(CACHE_FILE, ODS_FILE)                                           ## load the document which is processed to expand the spanned rows
    # # style_element = doc.styles.getElementsByType(Style)                                        ## Maybe no need for color, I can simply search for * or ●
    # expand_full_table = expand_column_repeated(expand_row_table)                                 ## full table expanded from spanned rows and spanned/repeated columns
    # duty_dates = expand_full_table[0][6:]                                                        ## dates for each day without get rid of vacations
    # duty_times = expand_full_table[2][6:]                                                        ## runs for each day without get rid of vacations
    # content_table = expand_full_table[84:654]                                                    ## content table is a table only contains contents of runs
    # # get_cell_from_index(145 - 85, column_letter_to_index("E"), content_table)                  ## print specific cell's content for content table
    # # get_cell_from_index(2, column_letter_to_index("F"), expand_row_table)                      ## print specific cell for different table
    # indices = get_content_index(content_table)                                                   ## indices of the cells in the whole table that has *, ●, dnf, or fr
    # warehouse = get_warehouse_statistic(content_table)                                           ## runs at each warehouse
    # plot_warehouse(warehouse)                                                                  ## plot the statistical graph of the runs for each warehouse
    # get_average_runs_for_dates("2022/6/6", "2022/8/31", duty_dates, duty_times)                  ## get the average runs between date1 and date2
    dtr_analyzer = Analyzer(CACHE_FILE, ODS_FILE)
    dtr_analyzer.get_average_runs_for_dates("2023/10/11", "2024/1/10")  
    dtr_analyzer.plot_warehouse()
    dtr_analyzer.plot_odor_trends()
    


main()

