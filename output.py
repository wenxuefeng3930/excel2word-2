import docxtpl as doc
import mailmerge


def create_doc(filename, data, template):
    document = mailmerge.MailMerge(template)
    # document.merge('A/CNO.账户号码'=data.get('A/CNO.账户号码'))
    data_list = data.get("list")
    obj = data_list[0]
    # print(obj)
    document.merge(obj)
    document.merge_rows("基金名称", data_list)
    document.write(filename)
    # tpl = doc.DocxTemplate(template)
    # tpl.render(obj)
    # tpl.save(filename)