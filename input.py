import openpyxl as xlrd


def line_2_arr(row):
    row_vals = []
    for cel in row:
        row_vals.append(str(cel.value).replace(" ", ""))
    return row_vals


def generate_data(row, keys, data_list, key):
    data = {}
    data_arr = []
    sub = {}
    i = 0
    for cel in row:
        cur_key = keys[i]
        if cur_key == key:
            data[key] = cel.value
        else:
            sub[keys[i]] = cel.value
        i += 1
    data_arr.append(sub)

    data_key = str(data.get(key))
    old_data = data_list.get(data_key)
    if old_data:
        old_list = old_data["list"]
        if old_list:
            pass
        for old_sub in old_list:
            data_arr.append(old_sub)

    data["list"] = data_arr
    return data


def export_data(template_name, key):
    wd = xlrd.load_workbook("./input/"+template_name)
    sheets = wd.worksheets
    sheet = sheets[0]
    rows = sheet.rows
    keys = []
    data_list = {}
    count = 1
    for row in rows:
        if count == 1:
            keys = line_2_arr(row)
        else:
            data = generate_data(row, keys, data_list, key)
            data_list[str(data.get(key))] = data
        count += 1
    return data_list
