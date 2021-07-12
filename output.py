# import docxtpl as doc
import mailmerge


def create_doc(filename, data, template):
    # tpl = doc.DocxTemplate(template)
    # tpl.render(data)
    # tpl.save(filename)
    document = mailmerge.MailMerge(template)
    # document.merge(data)
    data_list = data.get("list")
    obj = data_list[0]
    print(obj)
    document.merge(obj)
    # document.merge_rows(list(obj.keys())[0], data_list)
    document.write(filename)
