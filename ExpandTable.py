from odf.opendocument import load
from odf.table import Table, TableRow, TableCell

def find_gap(lst):
    lst.sort()
    for i in range(1, len(lst)):
        if lst[i] - lst[i - 1] > 1:
            return i
    return None

def load_data(file_path):
    doc = load(file_path)
    return doc

def expand_row_merged_cells(doc):
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
                    gap_index = find_gap(track_index[row_index + i])
                    if gap_index:
                        col_index = gap_index
                    else: 
                        if expanded_table[row_index + i][col_index] != "":
                            col_index = len(expanded_table[row_index + i])

                while len(expanded_table[row_index + i]) <= col_index:  # 確保列存在
                    expanded_table[row_index + i].append("")               
                expanded_table[row_index + i][col_index] = cell
            # if col_index not in track_index[row_index + i]:
                track_index[row_index + i].append(col_index)
            # else:
            #     print("Repeated Column index ?!!")
                
            col_index += 1
    
    return expanded_table
            
def expand_column_repeated(table):   
    expanded_table = [[] for _ in range(len(table))]
    for row in range(len(table)):
        for col, cell in enumerate(table[row]):
            try:
                col_span = cell.getAttribute("numbercolumnsspanned")
                col_span = int(col_span) if col_span is not None else 1
            
                col_repeat = cell.getAttribute("numbercolumnsrepeated")
                col_repeat = int(col_repeat) if col_repeat is not None else 1
            except AttributeError as e:
                
                print(f"Error {e} at index: ({(row, col)})")
            if  col_span > 1: 
                for _ in range(col_span):
                    expanded_table[row].append(cell)
                continue
            if col_repeat < 200:
                for _ in range(col_repeat):
                    if col_repeat == 1:
                        expanded_table[row].append(cell)
                    else:
                        expanded_table[row].append(TableCell())
    return expanded_table       
                
def get_row_value(row, table):
    for i in range(len(table[row])):
        print(f"Row {row} of Column index {i} for content:{str(table[row][i])}")
    
def get_cell_from_index(row, col, table):
    print(f"Index:{(row, col)}, Cell:{str(table[row][col])}")
