from ExpandTable import *
from CellProcess import *
from odf.style import Style

def main():
    doc = load_data("2023-3months-DTR.ods")
    expand_row_table= expand_row_merged_cells(doc)
    expand_full_table = expand_column_repeated(expand_row_table)
    # col_index = column_letter_to_index("J")
    # print(col_index)
    get_cell_from_index(193, column_letter_to_index("J"), expand_full_table)
    # get_row_value(156, expand_full_table)
    # for i, row in enumerate(expand_full_table):
    #     print(f"Length of Row {i}: {len(row)}")
    # print(len(expand_row_table[744]))
    # print(expand_row_table[0][-1].getAttribute("numbercolumnsrepeated"))    

    
main()