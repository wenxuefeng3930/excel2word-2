import docxtpl as doc
import mailmerge
from operator import attrgetter, itemgetter
from locale import setlocale, atof, LC_NUMERIC

setlocale(LC_NUMERIC, 'English_US')


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


def create_doc(filename, data, template, key, type):
    filename = filename.replace(" ", "_")
    print("正在生成文件：[" + filename + "]")
    document = mailmerge.MailMerge(template)
    data_list = data.get("list")
    total = 0
    if type == "template1":
        for data in data_list:
            if data.get("TOTAL_an_ke_hu_") is not None and is_number(atof(data.get("TOTAL_an_ke_hu_"))):
                total += atof(data.get("TOTAL_an_ke_hu_"))
            if data.get("cun_xu_fen_e") is not None and is_number(atof(data.get("cun_xu_fen_e"))):
                data["cun_xu_fen_e"] = format(atof(data.get("cun_xu_fen_e")), ",.4f")
            if data.get("qi_mo_dan_wei_jing_zhi") is not None and is_number(atof(data.get("qi_mo_dan_wei_jing_zhi"))):
                data["qi_mo_dan_wei_jing_zhi"] = format(atof(data.get("qi_mo_dan_wei_jing_zhi")), ",.4f")
            if data.get("qi_mo_zi_chan_jing_zhi") is not None and is_number(atof(data.get("qi_mo_zi_chan_jing_zhi"))):
                data["qi_mo_zi_chan_jing_zhi"] = format(atof(data.get("qi_mo_zi_chan_jing_zhi")), ",.2f")
            if data.get("jun_heng_bu_chang") is not None and is_number(atof(data.get("jun_heng_bu_chang"))):
                data["jun_heng_bu_chang"] = format(atof(data.get("jun_heng_bu_chang")), ",.2f")
            if data.get("qi_mo_dan_wei_jing_zhi__han_EQ") is not None and \
                    is_number(atof(data.get("qi_mo_dan_wei_jing_zhi__han_EQ"))):
                data["qi_mo_dan_wei_jing_zhi__han_EQ"] = format(atof(data.get("qi_mo_dan_wei_jing_zhi__han_EQ")), ",.4f")
            if data.get("qi_mo_zi_chan_jing_zhi_afterEQ") is not None and \
                    is_number(atof(data.get("qi_mo_zi_chan_jing_zhi_afterEQ"))):
                data["qi_mo_zi_chan_jing_zhi_afterEQ"] = format(atof(data.get("qi_mo_zi_chan_jing_zhi_afterEQ")), ",.2f")
        data_list = sorted(data_list, key=itemgetter('jiao_yi_kai_fang_ri'), reverse=False)
        document.merge_rows("ji_jin_ming_cheng", data_list)
        obj = data_list[0]
        document.merge(ACNO_zhang_hu_hao_ma=key,
                       StatementDate_jie_dan_ri_qi=obj.get('StatementDate_jie_dan_ri_qi'),
                       ke_hu_xing_ming=obj.get('ke_hu_xing_ming'),
                       ying_wen_xing_ming=obj.get('ying_wen_xing_ming'),
                       zhong_wen_di_zhi=obj.get('zhong_wen_di_zhi'),
                       ying_wen_di_zhi=obj.get('ying_wen_di_zhi'),
                       Email_dian_you_di_zhi=obj.get('Email_dian_you_di_zhi'),
                       TOTAL_an_ke_hu_=format(total, ",.2f"),
                       )
    elif type == "template2":
        data_list = sorted(data_list, key=itemgetter('jiao_yi_kai_fang_ri'), reverse=False)
        document.merge_rows("jiao_yi_kai_fang_ri", data_list)
        obj = data_list[0]
        document.merge(ke_hu_dai_ma=key,
                       jie_dan_ri_qi=obj.get('jie_dan_ri_qi'),
                       ke_hu_xing_ming=obj.get('ke_hu_xing_ming'),
                       ying_wen_xing_ming=obj.get('ying_wen_xing_ming'),
                       zhong_wen_di_zhi=obj.get('zhong_wen_di_zhi'),
                       ying_wen_di_zhi=obj.get('ying_wen_di_zhi'),
                       ke_hu_you_xiang=obj.get('ke_hu_you_xiang'),
                       )

    document.write(filename)
    print("生成文件：[" + filename + "]完成")
