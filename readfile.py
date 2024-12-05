import sys
import re
from odf.opendocument import load
from odf.table import Table, TableRow, TableCell
from odf.text import P
from odf.office import Annotation
from odf.style import Style, TableCellProperties
from odf.element import Element  # 通用解析元素

green = "#66cc00"

def load_data(file_path):
    doc = load(file_path)
    return doc

def get_cell_color(cell, style_elem):
    cell_style = cell.getAttribute("stylename")
    if cell_style:
        for style in style_elem:
            if style.getAttribute("name") == cell_style:
                props = style.getElementsByType(TableCellProperties)
                for prop in props:
                    bg_color = prop.getAttribute("backgroundcolor")
                    if bg_color:
                        return bg_color

def extract_text_from_element(element, layer=0):
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
            text += extract_text_from_element(child, layer)
    return text

def get_cell_content_and_annotation(cell):
    cell_content = ""
    annotation_text = ""
    for child in cell.childNodes:
        if child.tagName == "text:p":  # 確保是文本段落
            cell_content += "".join(text_node.data for text_node in child.childNodes if text_node.nodeType == 3)
        if child.tagName == "office:annotation":
            annotation_text = extract_text_from_element(child)
    return cell_content, annotation_text

def annotation_segment_process(text):
    pattern_start = r"^(Run\d+)(\d+\.)"
    match = re.match(pattern_start, text)
    segments = []
    if match:
        run_part = match.group(1)  # Run1
        first_item_number = match.group(2)  # 1.
        segments.append(run_part)
        segments.append(first_item_number + text[match.end():].split('2.')[0].strip())  # 添加Run1後的段落
        # 更新剩餘文本
        text = "2." + text.split('2.')[-1]
    # 2. 處理後續項次
    pattern_items = r"(\d+\.)"  # 匹配項次 2., 3., 4.
    parts = re.split(pattern_items, text)
    for i in range(1, len(parts), 2):
        item_number = parts[i].strip()
        content = parts[i + 1].strip() if i + 1 < len(parts) else ""
        segments.append(f"{item_number}{content}")

    return segments

def print_cell_and_annotation(cell):
    (content, annotation_text) = get_cell_content_and_annotation(cell)
    print(f"Cell Content: {content}")
    segs = annotation_segment_process(annotation_text)
    print("Annotation Content:\n")
    for segment in segs:
        print(f"- {segment}")


def main():
    doc = load_data("2023-3months-DTR.ods")
    style_elem = doc.automaticstyles.getElementsByType(Style)
    dtr_custy = doc.spreadsheet.childNodes[0]
    rows = dtr_custy.getElementsByType(TableRow)
    # cell_values, expanded_cells = expand_cells(row)
    index = 0
    cell_list = []
    annotation_list = []
    for cell in rows[144].getElementsByType(TableCell):
        cell_content, annotation_content = get_cell_content_and_annotation(cell)
        if cell_content:
            cell_list.append((cell_content, index))
        if annotation_content:
            annotation_list.append((annotation_content, index))
        index += 1
    for item in cell_list:
        print(item)
    for item in annotation_list:
        print(item)
    
# main()