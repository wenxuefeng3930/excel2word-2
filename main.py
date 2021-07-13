import input
import output
import configparser

# Press the green button in the gutter to run the script.
if __name__ == '__main__':
    # data_list = input.export_data(sys.argv[0])
    # for data in data_list:
    #     output.create_doc(sys.argv[1], data)
    config = configparser.ConfigParser()
    config.read("config.ini", encoding="utf-8")
    file = config.get("input", "excel")

    key = config.get("data", "key")

    path = config.get("output", "path")
    file_prefix = config.get("output", "prefix")
    name_key = config.get("output", "name-key")
    template = config.get("output", "template")

    data_list = input.export_data(file, key)
    print(data_list)
    for key in data_list:
        data = data_list.get(key)
        output.create_doc(path+file_prefix+"-"+str(key)+".docx", data, template, key)
