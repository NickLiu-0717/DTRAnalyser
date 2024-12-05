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

    for row_index, row in enumerate(sheet.getElementsByType(TableRow)):
        # 確保展開表格有足夠的行
        while len(expanded_table) <= row_index:
            expanded_table.append([])

        for col_index, cell in enumerate(row.getElementsByType(TableCell)):
            # 獲取儲存格值
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
            # row_span = cell.attributes.get(('urn:oasis:names:tc:opendocument:xmlns:table:1.0', 'number-rows-spanned'))
            row_span = int(row_span) if row_span is not None else 1
            if row_index == 144 and col_index == 6:
                print(cell.getAttribute("numbercolumnsrepeated"))
                # print(cell.attributes.get(('urn:oasis:names:tc:opendocument:xmlns:table:1.0', 'number-rows-spanned')))
                # print(row_span)
                # print(col_index)
            # 填充合併範圍的行
            for i in range(row_span):
                while len(expanded_table) <= row_index + i:  # 確保目標行存在
                    expanded_table.append([])
                while len(expanded_table[row_index + i]) <= col_index:  # 確保列存在
                    expanded_table[row_index + i].append("")
                expanded_table[row_index + i][col_index] = cell
    # print(get_cell_color(expanded_table[144][6], style_elem))
    return expanded_table

expanded_table = expand_row_merged_cells("2023-3months-DTR.ods")
# print(len(expanded_table[0]), len(expanded_table[1]))
# print(expanded_table[144][0])
# print(column_letter_to_index("AN"))