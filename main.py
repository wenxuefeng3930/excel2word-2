import sys

import input
import output
import configparser

# Press the green button in the gutter to run the script.
if __name__ == '__main__':
    # config_path = input.export_data(sys.argv[0])
    # for data in data_list:
    #     output.create_doc(sys.argv[1], data)
    type = sys.argv[1]

    if type == 'template1':
        config = configparser.ConfigParser()
        config.read("./config.ini", encoding="utf-8")
        file = config.get("input", "excel")
        in_path = config.get("input", "path")

        key = config.get("data", "key")

        path = config.get("output", "path")
        file_prefix = config.get("output", "prefix")
        name_key = config.get("output", "name-key")
        template = config.get("output", "template")

        data_list = input.export_data(file, key, in_path)
        for key in data_list:
            data = data_list.get(key)
            _list = data.get("list")
            name = str(_list[0].get("ke_hu_xing_ming"))
            product = str(_list[0].get("ji_jin_ming_cheng"))
            end = str(_list[0].get("StatementDate_jie_dan_ri_qi"))
            output.create_doc(path + name + "-" + product.replace("/", "-") + "-" + end + ".docx", data, template, key, type)
    elif type == 'template2':
        config = configparser.ConfigParser()
        config.read("./config.ini", encoding="utf-8")
        file = config.get("input_t2", "excel")
        in_path = config.get("input_t2", "path")

        key = config.get("data_t2", "key")

        path = config.get("output_t2", "path")
        file_prefix = config.get("output_t2", "prefix")
        name_key = config.get("output_t2", "name-key")
        template = config.get("output_t2", "template")

        data_list = input.export_data(file, key, in_path)
        for key in data_list:
            data = data_list.get(key)
            _list = data.get("list")
            name = str(_list[0].get("ke_hu_xing_ming"))
            product = str(_list[0].get("chan_pin_ming_cheng"))
            end = "成交结单"
            output.create_doc(path + product.replace("/", "-") + "-" + name + "-" + end + ".docx", data, template, key, type)


