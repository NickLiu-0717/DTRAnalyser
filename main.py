from ExpandTable import *
from CellProcess import *
from GrabContent import *
from odf.style import Style
import pandas as pd
import matplotlib.pyplot as plt
import matplotlib.font_manager as fm
from analyzer import *
from useful_func import *
# flattened_data = []

# def flatten_dict(d, keys=[]):   
#     for k, v in d.items():
#         if isinstance(v, dict):
#             flatten_dict(v, keys + [k])
#         else:
#             flattened_data.append(keys + [k, v])

# def plot_warehouse(warehouse):
#     font_path = "msjhl.ttc"
#     font_prop = fm.FontProperties(fname=font_path)
#     flatten_dict(warehouse)
#     columns = ["Warehouse", "Subcategory", "Detail", "Value"]
#     df = pd.DataFrame(flattened_data, columns=columns)
#     summary = df.groupby(["Warehouse", "Subcategory"])["Value"].sum().unstack()
#     summary["Total"] = summary.sum(axis=1)  # Add a 'Total' column for sorting
#     summary = summary.sort_values("Total", ascending=False).drop(columns=["Total"])  # Sort and drop 'Total'
#     # Plot
#     summary.plot(kind="bar", stacked=True, figsize=(10, 6))
#     plt.title("Total Imports and Exports by Category")
#     plt.ylabel("Value")
#     plt.xlabel("Warehouse")
#     plt.xticks(fontproperties=font_prop)
#     plt.legend(title="Type", prop=font_prop)
#     plt.show()


def main():
    CACHE_FILE = "cached_ods.pkl"
    ODS_FILE = "2023-3months-DTR.ods"
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
    dtr_analyzer.run_all_step()  
    dtr_analyzer.get_average_runs_for_dates("2022/6/6", "2022/8/31")  
    


main()

