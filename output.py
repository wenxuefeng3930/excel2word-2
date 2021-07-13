import docxtpl as doc
import mailmerge


def is_number(s):
    try:
        float(s)
        return True
    except ValueError:
        pass

    try:
        import unicodedata
        unicodedata.numeric(s)
        return True
    except (TypeError, ValueError):
        pass

    return False


def create_doc(filename, data, template, key):
    document = mailmerge.MailMerge(template)
    # document.merge('A/CNO.账户号码'=data.get('A/CNO.账户号码'))
    data_list = data.get("list")
    total = 0
    for data in data_list:
        if data.get("TOTAL_an_ke_hu_") is not None and is_number(data.get("TOTAL_an_ke_hu_")):
            total += float(data.get("TOTAL_an_ke_hu_"))

    document.merge_rows("ji_jin_ming_cheng", data_list)
    obj = data_list[0]
    document.merge(ACNO_zhang_hu_hao_ma=key,
                   StatementDate_jie_dan_ri_qi=obj.get('StatementDate_jie_dan_ri_qi'),
                   ke_hu_xing_ming=obj.get('ke_hu_xing_ming'),
                   ying_wen_xing_ming=obj.get('ying_wen_xing_ming'),
                   zhong_wen_di_zhi=obj.get('zhong_wen_di_zhi'),
                   ying_wen_di_zhi=obj.get('ying_wen_di_zhi'),
                   Email_dian_you_di_zhi=obj.get('Email_dian_you_di_zhi'),
                   TOTAL_an_ke_hu_=str(total),
                   )

    # print(obj)
    # document.merge(obj)
    document.write(filename)
    # tpl = doc.DocxTemplate(template)
    # tpl.render(obj)
    # tpl.save(filename)
