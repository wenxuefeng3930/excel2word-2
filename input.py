import datetime
import openpyxl as xlrd
import xpinyin


def line_2_arr(row):
    p = xpinyin.Pinyin()
    row_vals = []
    for cel in row:
        cel_val = str(cel.value).replace(" ", "")
        cel_val = p.get_pinyin(cel_val, splitter="_").replace('(', '').replace(')', '').replace('（', '').replace('）', '').replace('.', '').replace('/', '')
        row_vals.append(cel_val)
    return row_vals


def generate_data(row, keys, data_list, key):
    p = xpinyin.Pinyin()
    data = {}
    data_arr = []
    sub = {}
    i = 0
    key = p.get_pinyin(key)
    for cel in row:
        cur_key = keys[i]
        if cur_key == key:
            data[key] = str(cel.value)
        else:
            if type(cel.value) == datetime.datetime:
                sub[keys[i]] = '{:%Y-%m-%d}'.format(cel.value)
            elif type(cel.value) == float:
                sub[keys[i]] = format(cel.value, ",.2f")
                # str(('%2f' % cel.value))
            else:
                sub[keys[i]] = str(cel.value)
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


def export_data(template_name, key, in_path):
    wd = xlrd.load_workbook(in_path + template_name, data_only=True)
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
