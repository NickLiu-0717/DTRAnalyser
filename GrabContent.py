from CellProcess import *

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

def get_warehouse_statistic(table): ##0 / 0 = 0
    warehouse = {}
    for row in table:
        if str(row[1]) not in warehouse:
            if str(row[4]) != "0 / 0 = 0":
                warehouse[str(row[1])] = {}
                warehouse[str(row[1])][str(row[2])] = {}
                num = int(str(row[4])[0])
                warehouse[str(row[1])][str(row[2])][str(row[3])] = num
        
        else:
            if str(row[2]) not in warehouse[str(row[1])]:
                if str(row[4]) != "0 / 0 = 0":
                    warehouse[str(row[1])][str(row[2])] = {}
                    num = int(str(row[4])[0])
                    warehouse[str(row[1])][str(row[2])][str(row[3])] = num
            else:
                if str(row[3]) not in warehouse[str(row[1])][str(row[2])]:
                    if row[4] != "0 / 0 = 0":
                        num = int(str(row[4])[0])
                        warehouse[str(row[1])][str(row[2])][str(row[3])] = num
    return warehouse

def get_runs(duty_times):
    runs_dict = {}
    for i, run in enumerate(duty_times):
        if str(run):
            if len(str(run)) > 1:
                if str(run)[0].isdigit() and str(run)[1] == "(":
                    runs_dict[i] = 2 * int(str(run)[0])
            else:
                if str(run)[0].isdigit():
                    runs_dict[i] = int(str(run)[0])
    return runs_dict

def get_average_runs_for_dates(date1, date2, duty_date, duty_times):
    str_duty_date = [str(date) for date in duty_date]
    if date1 in str_duty_date and date2 in str_duty_date:
        index1 = str_duty_date.index(date1)
        index2 = str_duty_date.index(date2)
    else:
        print("Input dates are not in the list")
    runs_dict = get_runs(duty_times)
    key = index1
    count_days = 0
    sum_runs = 0
    while key >= index1 and key <= index2:
        if key in runs_dict:
            sum_runs += runs_dict[key]
            count_days += 1
            key += 1
        else:
            key += 1
    print(f"The average runs between {date1} and {date2} is {round(sum_runs / count_days, 2)}")
    
# def get_trend_of_odor(indice, table):
    
        





    