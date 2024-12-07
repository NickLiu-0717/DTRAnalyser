from CellProcess import *


def map_func(cell, train_index, op_index, dnf_index, fr_index, row, col):
    c, a = get_cell_content_and_annotation(cell)
    if c == "*":
        train_index.append(row, col)
    if c == "●":
        op_index.append(row, col)
    if c == "DNF":
        if not a:
            dnf_index.append(row, col)
    if c == "FR":
        fr_index.append(row, col)
    return train_index, op_index, dnf_index, fr_index

def filter_func(cell):
    c, a = get_cell_content_and_annotation(cell)
    if c == "*":
        return "train"
    if c == "●":
        return "OP"
    if c == "DNF":
        if not a:
            return "DNF"
    if c == "FR":
        return "FR"

def get_content_index(table):
    indices = {"train": {}, "OP": {}, "DNF": {}, "FR": {}}
    for r, row in enumerate(table):
        for i, cell in enumerate(row):
            category = filter_func(cell)
            if category:
                if r not in indices[category]:
                    indices[category][r] = []
                indices[category][r].append(i)
    return indices
    
    