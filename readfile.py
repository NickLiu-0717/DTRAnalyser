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

def get_cell_color(cell):
    cell_style = cell.getAttribute("stylename")
    if cell_style:
        style_elem = doc.automaticstyles.getElementsByType(Style)
        for style in style_elem:
            if style.getAttribute("name") == cell_style:
                props = style.getElementsByType(TableCellProperties)
                for prop in props:
                    bg_color = prop.getAttribute("backgroundcolor")
                    if bg_color:
                        return bg_color

def extract_text_from_element(element, layer=0):
    """
    遞迴提取節點內的所有文本內容。
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
    content = ""
    annotation_text = ""
    for child in cell.childNodes:
        print(child.tagName)
        if child.tagName == "text:p":  # 確保是文本段落
            content += "".join(text_node.data for text_node in child.childNodes if text_node.nodeType == 3)
        if child.tagName == "office:annotation":
            annotation_text = extract_text_from_element(child)
    return (content, annotation_text)

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

def expand_cells(row):
    """
    展開行內的所有儲存格，考慮 table:number-columns-repeated 屬性，使用 tagName 判斷節點類型。
    """
    expanded_cells = []
    cell_values = {}
    for cell in row.childNodes:
        # 如果 cell 是元組，取第一個元素
        if isinstance(cell, tuple):
            cell = cell[0]

        # 確認 cell 是否有 tagName 和 getAttribute 方法
        if hasattr(cell, "tagName") and hasattr(cell, "getAttribute"):
            if cell.tagName in ["table:table-cell", "table:covered-table-cell"]:
                # 防禦性處理重複屬性
                try:
                    repeat = cell.getAttribute("table:number-columns-repeated")
                    repeat = int(repeat) if repeat else 1
                except (ValueError, TypeError):
                    repeat = 1
                expanded_cells.extend([cell] * repeat)
                for _ in range(repeat):
                    expanded_cells.append(cell)
                    # 如果有內容，記錄索引和內容
                    if hasattr(cell, "textContent") and cell.textContent and cell.textContent.strip():
                        cell_values[len(expanded_cells) - 1] = cell.textContent.strip()
            else:
                print(f"Skipped unsupported node: {cell.tagName}")
        else:
            print(f"Skipped invalid node: {cell}")
    return cell_values, expanded_cells



def main():
    doc = load_data("2023-3months-DTR.ods")
    dtr_custy = doc.spreadsheet.childNodes[0]
    row = dtr_custy.getElementsByType(TableRow)[145]
    cell_values, expanded_cells = expand_cells(row)
    cell = expanded_cells[5]
        # 提取儲存格內容

    # print(row.toXml(0, sys.stdout))
    # content = ""
    # count_child = []
    # for child in cell.childNodes:
    #     if child.tagName == "text:p":  # 確保是文本段落
    #         content += "".join(text_node.data for text_node in child.childNodes if text_node.nodeType == 3)
    #     if child.tagName == "office:annotation":
    #         annotation_text, count_child = extract_text_from_element(child)

    # print(count_child)
    (content, annotation_text) = get_cell_content_and_annotation(cell)
    print(f"Cell Content: {content}")
    segs = annotation_segment_process(annotation_text)
    for segment in segs:
        print(f"- {segment}")
    
    
main()