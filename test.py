import sys
import re
from odf.opendocument import load
from odf.table import Table, TableRow, TableCell
from odf.text import P
from odf.office import Annotation
from odf.style import Style, TableCellProperties
from odf.element import Element
from readfile import *

def column_letter_to_index(column_letter):
    """
    將英文字母列標記 (A, B, ..., Z, AA, AB, ...) 轉換為列索引 (1, 2, ..., 26, 27, ...)
    """
    column_letter = column_letter.upper()  # 確保是大寫
    index = 0
    for char in column_letter:
        index = index * 26 + (ord(char) - ord('A') + 1)
    return index

def expand_row_merged_cells(ods_file):
    # 載入 ODS 文件
    doc = load(ods_file)
    sheet = doc.spreadsheet.getElementsByType(Table)[0]
    style_elem = doc.automaticstyles.getElementsByType(Style)

    expanded_table = []  # 用於存放展開後的表格
    index_table = {}
    for row_index, row in enumerate(sheet.getElementsByType(TableRow)):
        # 確保展開表格有足夠的行
        while len(expanded_table) <= row_index:
            expanded_table.append([])
        col_index = 0
        for coin, cell in enumerate(row.getElementsByType(TableCell)):
            # 獲取儲存格值
            index_table[row_index] = coin
            cell_value = ""
            for child in cell.childNodes:
                if child.tagName == "text:p":
                    cell_value = "".join(text_node.data for text_node in child.childNodes if text_node.nodeType == 3)
                    # if row_index == 144 and col_index == 0:
                    #     print(cell_value)
            if cell_value is None:
                cell_value = ""
            
            # 獲取合併範圍
            row_span = cell.getAttribute("numberrowsspanned")
            row_span = int(row_span) if row_span is not None else 1
            # if row_index == 1 and col_index == 0:        ##Print出我想要的特定cell的Attribute
            #     # print(cell.attributes.items(), row_span)
            #     cell = TableCell()
            #     text = P()
            #     text.addText("test message")
            #     cell.addElement(text)
                
                

            # print((row_index, col_index), col_repeat)
                # print(row_index, col_index)
            # new_cell = TableCell()
                # print(row_span)
                # print(col_index)
            # 填充合併範圍的行         
            
            for i in range(row_span):
                while len(expanded_table) < row_index + row_span:  # 確保目標行存在
                    expanded_table.append([])
                if col_index < len(expanded_table[row_index + i]):
                    if expanded_table[row_index + i][col_index] != "":
                        col_index = len(expanded_table[row_index + i])
                while len(expanded_table[row_index + i]) <= col_index:  # 確保列存在
                    expanded_table[row_index + i].append("")             
                expanded_table[row_index + i][col_index] = cell
            col_index += 1
    
    return expanded_table, index_table
            
def expand_column_repeated(table):   
    expanded_table = [[] for _ in range(len(table))]
    for row in range(len(table)):
        for cell in table[row]:
            col_repeat = cell.getAttribute("numbercolumnsrepeated")
            col_repeat = int(col_repeat) if col_repeat is not None else 1
            # print((row, col_repeat))
            for _ in range(col_repeat):
                if col_repeat == 1:
                    expanded_table[row].append(cell)
                else:
                    expanded_table[row].append(TableCell())
    return expanded_table       
                

expanded_row_table, index_table = expand_row_merged_cells("test3.ods")
expanded_full_table = expand_column_repeated(expanded_row_table)
# for i in range(len(expanded_row_table)):
#     print(f"len of row difference {i} : {len(expanded_row_table[i]) - index_table[i] - 1}")
# print(str(expanded_full_table[1][0]))
# row = 8
# print(len(expanded_row_table[row]))

def get_row_value(row, table):
    for i in range(len(table[row])):
        print(str(table[row][i]), i)
    
def get_cell_from_index(row, col, table):
    print(f"Index:{(row, col)}, Cell:{str(table[row][col])}")
    
get_cell_from_index(31, 2, expanded_full_table)
# get_row_value(4, expanded_row_table)
# row = 4
# for row in range(len(expanded_row_table)):
#     for i in range(len(expanded_row_table[row])):
#         # print(str(expanded_row_table[row][i]), i)
#         if expanded_row_table[row][i] == "":
#             print(True, (row, i))

# print(type(expanded_table[2][0]))
# print(expanded_row_table[4][3].attributes)
# print(len(expanded_table[0]), len(expanded_table[1]))
# print(expanded_table[144][0])
# print(column_letter_to_index("AN"))
# for row in index_table:
#     print(f"len of true row {row} : {index_table[row] + 1}")