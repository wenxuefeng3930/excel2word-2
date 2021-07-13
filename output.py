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
    obj = data_list[0]
    document.merge(ACNO_zhang_hu_hao_ma=key,
                   StatementDate_jie_dan_ri_qi=obj.get('StatementDate_jie_dan_ri_qi'),
                   ke_hu_xing_ming=obj.get('ke_hu_xing_ming'),
                   ying_wen_xing_ming=obj.get('ying_wen_xing_ming'),
                   zhong_wen_di_zhi=obj.get('zhong_wen_di_zhi'),
                   ying_wen_di_zhi=obj.get('ying_wen_di_zhi'),
                   Email_dian_you_di_zhi=obj.get('Email_dian_you_di_zhi'),
                   ji_jin_ming_cheng=obj.get('ji_jin_ming_cheng'),
                   jiao_yi_lei_xing=obj.get('jiao_yi_lei_xing'),
                   jiao_yi_kai_fang_ri=obj.get('jiao_yi_kai_fang_ri'),
                   cun_xu_fen_e=obj.get('cun_xu_fen_e'),
                   qi_mo_dan_wei_jing_zhi=obj.get('qi_mo_dan_wei_jing_zhi'),
                   qi_mo_zi_chan_jing_zhi=obj.get('qi_mo_zi_chan_jing_zhi'),
                   jun_heng_bu_chang=obj.get('jun_heng_bu_chang'),
                   qi_mo_dan_wei_jing_zhi__han_EQ=obj.get('qi_mo_dan_wei_jing_zhi__han_EQ'),
                   qi_mo_zi_chan_jing_zhi_afterEQ=obj.get('qi_mo_zi_chan_jing_zhi_afterEQ'),
                   TOTAL_an_ke_hu_=obj.get('TOTAL_an_ke_hu_'),
                   TOTAL_gong_shi_=obj.get('TOTAL_gong_shi_'),
                   liang_ge_qi_mo_zi_chan_he_dui=obj.get('liang_ge_qi_mo_zi_chan_he_dui'),
                   qi_mo_dan_wei_jing_zhi_EQ_he_dui=obj.get('qi_mo_dan_wei_jing_zhi_EQ_he_dui'))

    # print(obj)
    # document.merge(obj)
    document.merge_rows("ji_jin_ming_cheng", data_list)
    document.write(filename)
    # tpl = doc.DocxTemplate(template)
    # tpl.render(obj)
    # tpl.save(filename)
