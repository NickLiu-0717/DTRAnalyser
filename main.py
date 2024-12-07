from ExpandTable import *
from CellProcess import *
from GrabContent import *
from odf.style import Style

def main():
    doc = load_data("2023-3months-DTR.ods")
    # style_element = doc.styles.getElementsByType(Style)   ## Maybe no need for color, I can simply search for * or ●
    expand_row_table= expand_row_merged_cells(doc)
    expand_full_table = expand_column_repeated(expand_row_table)
    expand_full_table = expand_full_table[:653]
    # cell = expand_full_table[144][6]
    # get_cell_from_index(144, column_letter_to_index("G"), expand_full_table)
    # c, a = get_cell_content_and_annotation(cell)
    indices = get_content_index(expand_full_table)
    print(indices["FR"])
   

    
main()