import pandas as pd
import json

df = pd.read_excel('blank_G02.xlsx')
print(f"Loaded {len(df)} rows")

# Annotations for rows 0-59 (from agent analysis)
annotations = {}

# Rows 0-29
annotations[0] = {"ann_related": 1, "ann_year": [2019], "ann_fin_flag": 0, "ann_fin_info": None, "third_party_flag": 0, "third_party_list": None}
annotations[1] = {"ann_related": 1, "ann_year": [2021], "ann_fin_flag": 0, "ann_fin_info": None, "third_party_flag": 0, "third_party_list": None}
annotations[2] = {"ann_related": 0, "ann_year": None, "ann_fin_flag": None, "ann_fin_info": None, "third_party_flag": None, "third_party_list": None}
annotations[3] = {"ann_related": 0, "ann_year": None, "ann_fin_flag": None, "ann_fin_info": None, "third_party_flag": None, "third_party_list": None}
annotations[4] = {"ann_related": 1, "ann_year": [2002], "ann_fin_flag": 1, "ann_fin_info": [{"year": 2002, "elements": ["利润"]}], "third_party_flag": 0, "third_party_list": None}
annotations[5] = {"ann_related": 1, "ann_year": [2021], "ann_fin_flag": 0, "ann_fin_info": None, "third_party_flag": 0, "third_party_list": None}
annotations[6] = {"ann_related": 1, "ann_year": [2019, 2020], "ann_fin_flag": 0, "ann_fin_info": None, "third_party_flag": 0, "third_party_list": None}
annotations[7] = {"ann_related": 0, "ann_year": None, "ann_fin_flag": None, "ann_fin_info": None, "third_party_flag": None, "third_party_list": None}
annotations[8] = {"ann_related": 1, "ann_year": [2018, 2019, 2020], "ann_fin_flag": 1, "ann_fin_info": [{"year": 2018, "elements": ["资产", "负债", "所有者权益"]}, {"year": 2019, "elements": ["资产", "负债", "所有者权益"]}, {"year": 2020, "elements": ["资产", "负债", "所有者权益"]}], "third_party_flag": 0, "third_party_list": None}
annotations[9] = {"ann_related": 1, "ann_year": [2019, 2020, 2021, 2022], "ann_fin_flag": 0, "ann_fin_info": None, "third_party_flag": 0, "third_party_list": None}
annotations[10] = {"ann_related": 0, "ann_year": None, "ann_fin_flag": None, "ann_fin_info": None, "third_party_flag": None, "third_party_list": None}
annotations[11] = {"ann_related": 0, "ann_year": None, "ann_fin_flag": None, "ann_fin_info": None, "third_party_flag": None, "third_party_list": None}
annotations[12] = {"ann_related": 1, "ann_year": [2016, 2017, 2018, 2019, 2020, 2021], "ann_fin_flag": 1, "ann_fin_info": [{"year": 2016, "elements": ["收入", "利润"]}, {"year": 2017, "elements": ["收入", "利润"]}, {"year": 2018, "elements": ["收入", "利润"]}, {"year": 2019, "elements": ["收入", "利润"]}, {"year": 2020, "elements": ["收入", "利润"]}, {"year": 2021, "elements": ["收入", "利润", "在建工程"]}], "third_party_flag": 0, "third_party_list": None}
annotations[13] = {"ann_related": 0, "ann_year": None, "ann_fin_flag": None, "ann_fin_info": None, "third_party_flag": None, "third_party_list": None}
annotations[14] = {"ann_related": 0, "ann_year": None, "ann_fin_flag": None, "ann_fin_info": None, "third_party_flag": None, "third_party_list": None}
annotations[15] = {"ann_related": 0, "ann_year": None, "ann_fin_flag": None, "ann_fin_info": None, "third_party_flag": None, "third_party_list": None}
annotations[16] = {"ann_related": 0, "ann_year": None, "ann_fin_flag": None, "ann_fin_info": None, "third_party_flag": None, "third_party_list": None}
annotations[17] = {"ann_related": 1, "ann_year": [2013, 2014, 2015, 2016], "ann_fin_flag": 1, "ann_fin_info": [{"year": 2013, "elements": ["收入", "利润"]}, {"year": 2014, "elements": ["收入", "利润"]}, {"year": 2015, "elements": ["收入", "利润"]}, {"year": 2016, "elements": ["收入", "利润"]}], "third_party_flag": 0, "third_party_list": None}
annotations[18] = {"ann_related": 0, "ann_year": None, "ann_fin_flag": None, "ann_fin_info": None, "third_party_flag": None, "third_party_list": None}
annotations[19] = {"ann_related": 0, "ann_year": None, "ann_fin_flag": None, "ann_fin_info": None, "third_party_flag": None, "third_party_list": None}
annotations[20] = {"ann_related": 1, "ann_year": [2008], "ann_fin_flag": 1, "ann_fin_info": [{"year": 2008, "elements": ["资产", "负债", "所有者权益"]}], "third_party_flag": 0, "third_party_list": None}
annotations[21] = {"ann_related": 0, "ann_year": None, "ann_fin_flag": None, "ann_fin_info": None, "third_party_flag": None, "third_party_list": None}
annotations[22] = {"ann_related": 0, "ann_year": None, "ann_fin_flag": None, "ann_fin_info": None, "third_party_flag": None, "third_party_list": None}
annotations[23] = {"ann_related": 1, "ann_year": [2004], "ann_fin_flag": 0, "ann_fin_info": None, "third_party_flag": 0, "third_party_list": None}
annotations[24] = {"ann_related": 0, "ann_year": None, "ann_fin_flag": None, "ann_fin_info": None, "third_party_flag": None, "third_party_list": None}
annotations[25] = {"ann_related": 1, "ann_year": [2015, 2016], "ann_fin_flag": 1, "ann_fin_info": [{"year": 2015, "elements": ["收入", "利润"]}, {"year": 2016, "elements": ["收入", "利润"]}], "third_party_flag": 0, "third_party_list": None}
annotations[26] = {"ann_related": 0, "ann_year": None, "ann_fin_flag": None, "ann_fin_info": None, "third_party_flag": None, "third_party_list": None}
annotations[27] = {"ann_related": 0, "ann_year": None, "ann_fin_flag": None, "ann_fin_info": None, "third_party_flag": None, "third_party_list": None}
annotations[28] = {"ann_related": 0, "ann_year": None, "ann_fin_flag": None, "ann_fin_info": None, "third_party_flag": None, "third_party_list": None}
annotations[29] = {"ann_related": 1, "ann_year": [2012], "ann_fin_flag": 1, "ann_fin_info": [{"year": 2012, "elements": ["利润", "收入"]}], "third_party_flag": 0, "third_party_list": None}

# Rows 30-59
annotations[30] = {"ann_related": 0, "ann_year": None, "ann_fin_flag": None, "ann_fin_info": None, "third_party_flag": None, "third_party_list": None}
annotations[31] = {"ann_related": 1, "ann_year": [2022], "ann_fin_flag": 1, "ann_fin_info": [{"year": 2022, "elements": ["资产"]}], "third_party_flag": 0, "third_party_list": None}
annotations[32] = {"ann_related": 0, "ann_year": None, "ann_fin_flag": None, "ann_fin_info": None, "third_party_flag": None, "third_party_list": None}
annotations[33] = {"ann_related": 0, "ann_year": None, "ann_fin_flag": None, "ann_fin_info": None, "third_party_flag": None, "third_party_list": None}
annotations[34] = {"ann_related": 0, "ann_year": None, "ann_fin_flag": None, "ann_fin_info": None, "third_party_flag": None, "third_party_list": None}
annotations[35] = {"ann_related": 0, "ann_year": None, "ann_fin_flag": None, "ann_fin_info": None, "third_party_flag": None, "third_party_list": None}
annotations[36] = {"ann_related": 1, "ann_year": [2015, 2016, 2017, 2018, 2019, 2020], "ann_fin_flag": 0, "ann_fin_info": None, "third_party_flag": 0, "third_party_list": None}
annotations[37] = {"ann_related": 0, "ann_year": None, "ann_fin_flag": None, "ann_fin_info": None, "third_party_flag": None, "third_party_list": None}
annotations[38] = {"ann_related": 0, "ann_year": None, "ann_fin_flag": None, "ann_fin_info": None, "third_party_flag": None, "third_party_list": None}
annotations[39] = {"ann_related": 0, "ann_year": None, "ann_fin_flag": None, "ann_fin_info": None, "third_party_flag": None, "third_party_list": None}
annotations[40] = {"ann_related": 1, "ann_year": [2014, 2015, 2016], "ann_fin_flag": 1, "ann_fin_info": [{"year": 2014, "elements": ["收入", "利润"]}, {"year": 2015, "elements": ["收入", "利润"]}, {"year": 2016, "elements": ["收入", "利润"]}], "third_party_flag": 0, "third_party_list": None}
annotations[41] = {"ann_related": 1, "ann_year": [2020], "ann_fin_flag": 0, "ann_fin_info": None, "third_party_flag": 0, "third_party_list": None}
annotations[42] = {"ann_related": 0, "ann_year": None, "ann_fin_flag": None, "ann_fin_info": None, "third_party_flag": None, "third_party_list": None}
annotations[43] = {"ann_related": 1, "ann_year": [2023], "ann_fin_flag": 0, "ann_fin_info": None, "third_party_flag": 0, "third_party_list": None}
annotations[44] = {"ann_related": 0, "ann_year": None, "ann_fin_flag": None, "ann_fin_info": None, "third_party_flag": None, "third_party_list": None}
annotations[45] = {"ann_related": 1, "ann_year": [2024], "ann_fin_flag": 0, "ann_fin_info": None, "third_party_flag": 0, "third_party_list": None}
annotations[46] = {"ann_related": 0, "ann_year": None, "ann_fin_flag": None, "ann_fin_info": None, "third_party_flag": None, "third_party_list": None}
annotations[47] = {"ann_related": 0, "ann_year": None, "ann_fin_flag": None, "ann_fin_info": None, "third_party_flag": None, "third_party_list": None}
annotations[48] = {"ann_related": 0, "ann_year": None, "ann_fin_flag": None, "ann_fin_info": None, "third_party_flag": None, "third_party_list": None}
annotations[49] = {"ann_related": 1, "ann_year": [2021], "ann_fin_flag": 1, "ann_fin_info": [{"year": 2021, "elements": ["收入", "费用"]}], "third_party_flag": 0, "third_party_list": None}
annotations[50] = {"ann_related": 1, "ann_year": [2010], "ann_fin_flag": 0, "ann_fin_info": None, "third_party_flag": 0, "third_party_list": None}
annotations[51] = {"ann_related": 0, "ann_year": None, "ann_fin_flag": None, "ann_fin_info": None, "third_party_flag": None, "third_party_list": None}
annotations[52] = {"ann_related": 0, "ann_year": None, "ann_fin_flag": None, "ann_fin_info": None, "third_party_flag": None, "third_party_list": None}
annotations[53] = {"ann_related": 1, "ann_year": [2013], "ann_fin_flag": 1, "ann_fin_info": [{"year": 2013, "elements": ["收入"]}], "third_party_flag": 1, "third_party_list": [{"name": "山东魏桥铝电有限公司", "type": "客户", "role": "配合提前确认收入,货物未实际出库"}, {"name": "山东阳谷华泰进出口有限公司", "type": "客户", "role": "配合提前确认收入,货物未实际出库"}]}
annotations[54] = {"ann_related": 0, "ann_year": None, "ann_fin_flag": None, "ann_fin_info": None, "third_party_flag": None, "third_party_list": None}
annotations[55] = {"ann_related": 0, "ann_year": None, "ann_fin_flag": None, "ann_fin_info": None, "third_party_flag": None, "third_party_list": None}
annotations[56] = {"ann_related": 1, "ann_year": [2013], "ann_fin_flag": 1, "ann_fin_info": [{"year": 2013, "elements": ["收入"]}], "third_party_flag": 1, "third_party_list": [{"name": "山东魏桥铝电有限公司", "type": "客户", "role": "配合提前确认收入,货物未实际出库"}, {"name": "山东阳谷华泰进出口有限公司", "type": "客户", "role": "配合提前确认收入,货物未实际出库"}]}
annotations[57] = {"ann_related": 0, "ann_year": None, "ann_fin_flag": None, "ann_fin_info": None, "third_party_flag": None, "third_party_list": None}
annotations[58] = {"ann_related": 0, "ann_year": None, "ann_fin_flag": None, "ann_fin_info": None, "third_party_flag": None, "third_party_list": None}
annotations[59] = {"ann_related": 0, "ann_year": None, "ann_fin_flag": None, "ann_fin_info": None, "third_party_flag": None, "third_party_list": None}

# Rows 60-99 (continuing analysis)
annotations[60] = {"ann_related": 0, "ann_year": None, "ann_fin_flag": None, "ann_fin_info": None, "third_party_flag": None, "third_party_list": None}  # 业绩预告不准确但年报数据本身正确，预告披露问题
annotations[61] = {"ann_related": 0, "ann_year": None, "ann_fin_flag": None, "ann_fin_info": None, "third_party_flag": None, "third_party_list": None}  # 司法冻结披露问题，非年报数据
annotations[62] = {"ann_related": 0, "ann_year": None, "ann_fin_flag": None, "ann_fin_info": None, "third_party_flag": None, "third_party_list": None}  # 会计估计变更披露问题
annotations[63] = {"ann_related": 0, "ann_year": None, "ann_fin_flag": None, "ann_fin_info": None, "third_party_flag": None, "third_party_list": None}  # 业绩预告披露不准确，年报本身数据需核实
annotations[64] = {"ann_related": 0, "ann_year": None, "ann_fin_flag": None, "ann_fin_info": None, "third_party_flag": None, "third_party_list": None}  # 业绩预告与审计数据差异，未及时修正预告
annotations[65] = {"ann_related": 1, "ann_year": [2017], "ann_fin_flag": 0, "ann_fin_info": None, "third_party_flag": 0, "third_party_list": None}  # 关联交易未披露，非财务数据错误
annotations[66] = {"ann_related": 0, "ann_year": None, "ann_fin_flag": None, "ann_fin_info": None, "third_party_flag": None, "third_party_list": None}  # 减持未披露，高管交易违规
annotations[67] = {"ann_related": 0, "ann_year": None, "ann_fin_flag": None, "ann_fin_info": None, "third_party_flag": None, "third_party_list": None}  # 违规担保未披露，非财务数据
annotations[68] = {"ann_related": 0, "ann_year": None, "ann_fin_flag": None, "ann_fin_info": None, "third_party_flag": None, "third_party_list": None}  # 券商内控问题
annotations[69] = {"ann_related": 0, "ann_year": None, "ann_fin_flag": None, "ann_fin_info": None, "third_party_flag": None, "third_party_list": None}  # 募集资金管理不规范，披露问题
annotations[70] = {"ann_related": 0, "ann_year": None, "ann_fin_flag": None, "ann_fin_info": None, "third_party_flag": None, "third_party_list": None}  # 股东权益变动未披露
annotations[71] = {"ann_related": 0, "ann_year": None, "ann_fin_flag": None, "ann_fin_info": None, "third_party_flag": None, "third_party_list": None}  # 收购方案披露问题
annotations[72] = {"ann_related": 0, "ann_year": None, "ann_fin_flag": None, "ann_fin_info": None, "third_party_flag": None, "third_party_list": None}  # 股权转让回购安排披露问题
annotations[73] = {"ann_related": 0, "ann_year": None, "ann_fin_flag": None, "ann_fin_info": None, "third_party_flag": None, "third_party_list": None}  # 上市时未披露关联关系，证券违规
annotations[74] = {"ann_related": 0, "ann_year": None, "ann_fin_flag": None, "ann_fin_info": None, "third_party_flag": None, "third_party_list": None}  # 股东被动减持披露
annotations[75] = {"ann_related": 0, "ann_year": None, "ann_fin_flag": None, "ann_fin_info": None, "third_party_flag": None, "third_party_list": None}  # 市场垄断处罚
annotations[76] = {"ann_related": 0, "ann_year": None, "ann_fin_flag": None, "ann_fin_info": None, "third_party_flag": None, "third_party_list": None}  # 杂志摆放权信息披露问题
annotations[77] = {"ann_related": 1, "ann_year": [2019, 2020, 2021], "ann_fin_flag": 0, "ann_fin_info": None, "third_party_flag": 0, "third_party_list": None}  # 审计程序存在缺陷，涉及年报
annotations[78] = {"ann_related": 0, "ann_year": None, "ann_fin_flag": None, "ann_fin_info": None, "third_party_flag": None, "third_party_list": None}  # 业绩预告修正
annotations[79] = {"ann_related": 0, "ann_year": None, "ann_fin_flag": None, "ann_fin_info": None, "third_party_flag": None, "third_party_list": None}  # 华锐风电虚假信息披露，人员责任认定
annotations[80] = {"ann_related": 1, "ann_year": [2009, 2010, 2011, 2012, 2013, 2014], "ann_fin_flag": 1, "ann_fin_info": [{"year": 2009, "elements": ["收入"]}, {"year": 2010, "elements": ["收入"]}, {"year": 2011, "elements": ["收入"]}, {"year": 2012, "elements": ["收入"]}, {"year": 2013, "elements": ["收入"]}, {"year": 2014, "elements": ["收入"]}], "third_party_flag": 0, "third_party_list": None}
annotations[81] = {"ann_related": 0, "ann_year": None, "ann_fin_flag": None, "ann_fin_info": None, "third_party_flag": None, "third_party_list": None}  # 股东违规转让
annotations[82] = {"ann_related": 0, "ann_year": None, "ann_fin_flag": None, "ann_fin_info": None, "third_party_flag": None, "third_party_list": None}  # 股东增持披露违规
annotations[83] = {"ann_related": 1, "ann_year": [2022], "ann_fin_flag": 0, "ann_fin_info": None, "third_party_flag": 0, "third_party_list": None}  # 关联方资金往来披露问题
annotations[84] = {"ann_related": 0, "ann_year": None, "ann_fin_flag": None, "ann_fin_info": None, "third_party_flag": None, "third_party_list": None}  # IPO保荐违规
annotations[85] = {"ann_related": 0, "ann_year": None, "ann_fin_flag": None, "ann_fin_info": None, "third_party_flag": None, "third_party_list": None}  # 募集资金归还逾期
annotations[86] = {"ann_related": 0, "ann_year": None, "ann_fin_flag": None, "ann_fin_info": None, "third_party_flag": None, "third_party_list": None}  # 短线交易
annotations[87] = {"ann_related": 0, "ann_year": None, "ann_fin_flag": None, "ann_fin_info": None, "third_party_flag": None, "third_party_list": None}  # 短线交易
annotations[88] = {"ann_related": 0, "ann_year": None, "ann_fin_flag": None, "ann_fin_info": None, "third_party_flag": None, "third_party_list": None}  # 董事未出席董事会
annotations[89] = {"ann_related": 0, "ann_year": None, "ann_fin_flag": None, "ann_fin_info": None, "third_party_flag": None, "third_party_list": None}  # 募集资金管理问题
annotations[90] = {"ann_related": 1, "ann_year": [2017, 2018, 2019, 2020, 2021, 2022], "ann_fin_flag": 1, "ann_fin_info": [{"year": 2017, "elements": ["收入", "利润"]}, {"year": 2018, "elements": ["收入", "利润"]}, {"year": 2019, "elements": ["收入", "利润"]}, {"year": 2020, "elements": ["收入", "利润"]}, {"year": 2021, "elements": ["收入", "利润"]}, {"year": 2022, "elements": ["收入", "利润"]}], "third_party_flag": 0, "third_party_list": None}  # 信唐普华虚构第三方业务，但第三方是配合方作为交易对手被动参与，非主动协助
annotations[91] = {"ann_related": 0, "ann_year": None, "ann_fin_flag": None, "ann_fin_info": None, "third_party_flag": None, "third_party_list": None}  # 评估机构问题
annotations[92] = {"ann_related": 0, "ann_year": None, "ann_fin_flag": None, "ann_fin_info": None, "third_party_flag": None, "third_party_list": None}  # 短线交易
annotations[93] = {"ann_related": 0, "ann_year": None, "ann_fin_flag": None, "ann_fin_info": None, "third_party_flag": None, "third_party_list": None}  # 仲裁裁定披露
annotations[94] = {"ann_related": 0, "ann_year": None, "ann_fin_flag": None, "ann_fin_info": None, "third_party_flag": None, "third_party_list": None}  # 重组承诺未履行
annotations[95] = {"ann_related": 0, "ann_year": None, "ann_fin_flag": None, "ann_fin_info": None, "third_party_flag": None, "third_party_list": None}  # 会计师轮换违规

# Rows 100-185 (continuing)
annotations[96] = {"ann_related": 0, "ann_year": None, "ann_fin_flag": None, "ann_fin_info": None, "third_party_flag": None, "third_party_list": None}  # ST贵人业绩预告
annotations[97] = {"ann_related": 0, "ann_year": None, "ann_fin_flag": None, "ann_fin_info": None, "third_party_flag": None, "third_party_list": None}  # 杰华特业绩快报
annotations[98] = {"ann_related": 0, "ann_year": None, "ann_fin_flag": None, "ann_fin_info": None, "third_party_flag": None, "third_party_list": None}  # 内幕交易
annotations[99] = {"ann_related": 0, "ann_year": None, "ann_fin_flag": None, "ann_fin_info": None, "third_party_flag": None, "third_party_list": None}  # 退市风险披露问题

annotations[100] = {"ann_related": 0, "ann_year": None, "ann_fin_flag": None, "ann_fin_info": None, "third_party_flag": None, "third_party_list": None}  # 少缴税款，税务违规
annotations[101] = {"ann_related": 0, "ann_year": None, "ann_fin_flag": None, "ann_fin_info": None, "third_party_flag": None, "third_party_list": None}  # 短线交易
annotations[102] = {"ann_related": 0, "ann_year": None, "ann_fin_flag": None, "ann_fin_info": None, "third_party_flag": None, "third_party_list": None}  # ST贵人业绩预告
annotations[103] = {"ann_related": 0, "ann_year": None, "ann_fin_flag": None, "ann_fin_info": None, "third_party_flag": None, "third_party_list": None}  # 股权变更披露
annotations[104] = {"ann_related": 1, "ann_year": [2023], "ann_fin_flag": 0, "ann_fin_info": None, "third_party_flag": 0, "third_party_list": None}  # 资金占用，年报中有专项审计
annotations[105] = {"ann_related": 0, "ann_year": None, "ann_fin_flag": None, "ann_fin_info": None, "third_party_flag": None, "third_party_list": None}  # 业绩预告修正
annotations[106] = {"ann_related": 0, "ann_year": None, "ann_fin_flag": None, "ann_fin_info": None, "third_party_flag": None, "third_party_list": None}  # 康美药业业绩预告
annotations[107] = {"ann_related": 1, "ann_year": [2020], "ann_fin_flag": 1, "ann_fin_info": [{"year": 2020, "elements": ["收入"]}], "third_party_flag": 0, "third_party_list": None}  # 总额法核算导致收入披露不准确，影响收入确认
annotations[108] = {"ann_related": 0, "ann_year": None, "ann_fin_flag": None, "ann_fin_info": None, "third_party_flag": None, "third_party_list": None}  # 股东减持违规
annotations[109] = {"ann_related": 1, "ann_year": [2024], "ann_fin_flag": 1, "ann_fin_info": [{"year": 2024, "elements": ["资产"]}], "third_party_flag": 0, "third_party_list": None}  # 存货会计处理不规范，影响财务数据
annotations[110] = {"ann_related": 0, "ann_year": None, "ann_fin_flag": None, "ann_fin_info": None, "third_party_flag": None, "third_party_list": None}  # 药品销售违规
annotations[111] = {"ann_related": 1, "ann_year": [2020], "ann_fin_flag": 0, "ann_fin_info": None, "third_party_flag": 0, "third_party_list": None}  # 关联交易未及时审议披露
annotations[112] = {"ann_related": 1, "ann_year": [2022], "ann_fin_flag": 1, "ann_fin_info": [{"year": 2022, "elements": ["资产"]}], "third_party_flag": 0, "third_party_list": None}  # 关联方非经营性资金占用
annotations[113] = {"ann_related": 1, "ann_year": [2002, 2003, 2004], "ann_fin_flag": 0, "ann_fin_info": None, "third_party_flag": 0, "third_party_list": None}  # 向控股股东提供资金未履行审议程序和信息披露，但非年报财务数据错误
annotations[114] = {"ann_related": 0, "ann_year": None, "ann_fin_flag": None, "ann_fin_info": None, "third_party_flag": None, "third_party_list": None}  # 券商信息系统问题
annotations[115] = {"ann_related": 0, "ann_year": None, "ann_fin_flag": None, "ann_fin_info": None, "third_party_flag": None, "third_party_list": None}  # 审计机构问题
annotations[116] = {"ann_related": 1, "ann_year": [2005, 2006, 2007, 2008], "ann_fin_flag": 1, "ann_fin_info": [{"year": 2005, "elements": ["负债"]}, {"year": 2006, "elements": ["负债"]}, {"year": 2007, "elements": ["负债"]}, {"year": 2008, "elements": ["负债"]}], "third_party_flag": 0, "third_party_list": None}  # 贷款利息少计罚息，虚增利润
annotations[117] = {"ann_related": 1, "ann_year": [2004], "ann_fin_flag": 1, "ann_fin_info": [{"year": 2004, "elements": ["利润"]}], "third_party_flag": 0, "third_party_list": None}  # 少提坏账、虚构交易、少计利息导致年报虚假记载
annotations[118] = {"ann_related": 0, "ann_year": None, "ann_fin_flag": None, "ann_fin_info": None, "third_party_flag": None, "third_party_list": None}  # 重组披露超时
annotations[119] = {"ann_related": 0, "ann_year": None, "ann_fin_flag": None, "ann_fin_info": None, "third_party_flag": None, "third_party_list": None}  # 持股变动披露违规
annotations[120] = {"ann_related": 0, "ann_year": None, "ann_fin_flag": None, "ann_fin_info": None, "third_party_flag": None, "third_party_list": None}  # 可转债转股价格修正未披露
annotations[121] = {"ann_related": 1, "ann_year": [2003, 2004], "ann_fin_flag": 0, "ann_fin_info": None, "third_party_flag": 0, "third_party_list": None}  # 资金提供和担保未披露，非财务数据
annotations[122] = {"ann_related": 0, "ann_year": None, "ann_fin_flag": None, "ann_fin_info": None, "third_party_flag": None, "third_party_list": None}  # 董事买卖股票
annotations[123] = {"ann_related": 1, "ann_year": [2023], "ann_fin_flag": 0, "ann_fin_info": None, "third_party_flag": 0, "third_party_list": None}  # 非经营性资金往来未披露
annotations[124] = {"ann_related": 0, "ann_year": None, "ann_fin_flag": None, "ann_fin_info": None, "third_party_flag": None, "third_party_list": None}  # 股东减持未披露
annotations[125] = {"ann_related": 0, "ann_year": None, "ann_fin_flag": None, "ann_fin_info": None, "third_party_flag": None, "third_party_list": None}  # 非法利用他人账户
annotations[126] = {"ann_related": 1, "ann_year": [2014], "ann_fin_flag": 1, "ann_fin_info": [{"year": 2014, "elements": ["收入", "利润"]}], "third_party_flag": 0, "third_party_list": None}  # 跨期确认收入，虚增收入和利润
annotations[127] = {"ann_related": 0, "ann_year": None, "ann_fin_flag": None, "ann_fin_info": None, "third_party_flag": None, "third_party_list": None}  # 控股股东资金占用
annotations[128] = {"ann_related": 0, "ann_year": None, "ann_fin_flag": None, "ann_fin_info": None, "third_party_flag": None, "third_party_list": None}  # 股东买卖股票
annotations[129] = {"ann_related": 0, "ann_year": None, "ann_fin_flag": None, "ann_fin_info": None, "third_party_flag": None, "third_party_list": None}  # 内控缺陷
annotations[130] = {"ann_related": 0, "ann_year": None, "ann_fin_flag": None, "ann_fin_info": None, "third_party_flag": None, "third_party_list": None}  # 会计差错更正
annotations[131] = {"ann_related": 0, "ann_year": None, "ann_fin_flag": None, "ann_fin_info": None, "third_party_flag": None, "third_party_list": None}  # 权益变动披露
annotations[132] = {"ann_related": 0, "ann_year": None, "ann_fin_flag": None, "ann_fin_info": None, "third_party_flag": None, "third_party_list": None}  # 资产减值未及时审议
annotations[133] = {"ann_related": 0, "ann_year": None, "ann_fin_flag": None, "ann_fin_info": None, "third_party_flag": None, "third_party_list": None}  # 权益变动违规
annotations[134] = {"ann_related": 0, "ann_year": None, "ann_fin_flag": None, "ann_fin_info": None, "third_party_flag": None, "third_party_list": None}  # 高管窗口期减持
annotations[135] = {"ann_related": 0, "ann_year": None, "ann_fin_flag": None, "ann_fin_info": None, "third_party_flag": None, "third_party_list": None}  # 房产出售未及时披露
annotations[136] = {"ann_related": 0, "ann_year": None, "ann_fin_flag": None, "ann_fin_info": None, "third_party_flag": None, "third_party_list": None}  # EMPTY
annotations[137] = {"ann_related": 0, "ann_year": None, "ann_fin_flag": None, "ann_fin_info": None, "third_party_flag": None, "third_party_list": None}  # 口罩生产信息披露不完整
annotations[138] = {"ann_related": 0, "ann_year": None, "ann_fin_flag": None, "ann_fin_info": None, "third_party_flag": None, "third_party_list": None}  # 评估机构问题
annotations[139] = {"ann_related": 0, "ann_year": None, "ann_fin_flag": None, "ann_fin_info": None, "third_party_flag": None, "third_party_list": None}  # 营业部合规问题
annotations[140] = {"ann_related": 1, "ann_year": [2020, 2021], "ann_fin_flag": 0, "ann_fin_info": None, "third_party_flag": 0, "third_party_list": None}  # 佳通轮胎关联关系披露问题
annotations[141] = {"ann_related": 0, "ann_year": None, "ann_fin_flag": None, "ann_fin_info": None, "third_party_flag": None, "third_party_list": None}  # 监事会主席失联
annotations[142] = {"ann_related": 0, "ann_year": None, "ann_fin_flag": None, "ann_fin_info": None, "third_party_flag": None, "third_party_list": None}  # 券商内控
annotations[143] = {"ann_related": 0, "ann_year": None, "ann_fin_flag": None, "ann_fin_info": None, "third_party_flag": None, "third_party_list": None}  # 权益变动报告披露不完整
annotations[144] = {"ann_related": 0, "ann_year": None, "ann_fin_flag": None, "ann_fin_info": None, "third_party_flag": None, "third_party_list": None}  # 业绩预告修正
annotations[145] = {"ann_related": 0, "ann_year": None, "ann_fin_flag": None, "ann_fin_info": None, "third_party_flag": None, "third_party_list": None}  # 安全生产违规
annotations[146] = {"ann_related": 0, "ann_year": None, "ann_fin_flag": None, "ann_fin_info": None, "third_party_flag": None, "third_party_list": None}  # 关联交易未披露
annotations[147] = {"ann_related": 0, "ann_year": None, "ann_fin_flag": None, "ann_fin_info": None, "third_party_flag": None, "third_party_list": None}  # 担保未披露
annotations[148] = {"ann_related": 0, "ann_year": None, "ann_fin_flag": None, "ann_fin_info": None, "third_party_flag": None, "third_party_list": None}  # 券商两融违规
annotations[149] = {"ann_related": 0, "ann_year": None, "ann_fin_flag": None, "ann_fin_info": None, "third_party_flag": None, "third_party_list": None}  # 关联交易未披露
annotations[150] = {"ann_related": 0, "ann_year": None, "ann_fin_flag": None, "ann_fin_info": None, "third_party_flag": None, "third_party_list": None}  # 募集资金管理超期
annotations[151] = {"ann_related": 0, "ann_year": None, "ann_fin_flag": None, "ann_fin_info": None, "third_party_flag": None, "third_party_list": None}  # 持股稀释未披露
annotations[152] = {"ann_related": 0, "ann_year": None, "ann_fin_flag": None, "ann_fin_info": None, "third_party_flag": None, "third_party_list": None}  # 财务内控问题
annotations[153] = {"ann_related": 1, "ann_year": [2002, 2003, 2004], "ann_fin_flag": 1, "ann_fin_info": [{"year": 2002, "elements": ["收入", "利润"]}, {"year": 2003, "elements": ["收入", "利润"]}, {"year": 2004, "elements": ["收入", "利润"]}], "third_party_flag": 0, "third_party_list": None}
annotations[154] = {"ann_related": 1, "ann_year": [2018, 2019, 2020, 2021], "ann_fin_flag": 1, "ann_fin_info": [{"year": 2018, "elements": ["收入", "利润"]}, {"year": 2019, "elements": ["收入", "利润"]}, {"year": 2020, "elements": ["收入", "利润"]}, {"year": 2021, "elements": ["收入", "利润"]}], "third_party_flag": 1, "third_party_list": [{"name": "隋某力控制或指定的企业", "type": "其他企业", "role": "作为专网通信虚假自循环业务上下游主体签订合同并参与资金链条"}]}
annotations[155] = {"ann_related": 0, "ann_year": None, "ann_fin_flag": None, "ann_fin_info": None, "third_party_flag": None, "third_party_list": None}  # 高管减持
annotations[156] = {"ann_related": 0, "ann_year": None, "ann_fin_flag": None, "ann_fin_info": None, "third_party_flag": None, "third_party_list": None}  # 短线交易
annotations[157] = {"ann_related": 0, "ann_year": None, "ann_fin_flag": None, "ann_fin_info": None, "third_party_flag": None, "third_party_list": None}  # 门诊部违规
annotations[158] = {"ann_related": 0, "ann_year": None, "ann_fin_flag": None, "ann_fin_info": None, "third_party_flag": None, "third_party_list": None}  # 募投项目变更未审议
annotations[159] = {"ann_related": 0, "ann_year": None, "ann_fin_flag": None, "ann_fin_info": None, "third_party_flag": None, "third_party_list": None}  # 高管短线交易
annotations[160] = {"ann_related": 0, "ann_year": None, "ann_fin_flag": None, "ann_fin_info": None, "third_party_flag": None, "third_party_list": None}  # 评估机构问题
annotations[161] = {"ann_related": 0, "ann_year": None, "ann_fin_flag": None, "ann_fin_info": None, "third_party_flag": None, "third_party_list": None}  # 独董兼职超限
annotations[162] = {"ann_related": 0, "ann_year": None, "ann_fin_flag": None, "ann_fin_info": None, "third_party_flag": None, "third_party_list": None}  # 业绩预告差异
annotations[163] = {"ann_related": 0, "ann_year": None, "ann_fin_flag": None, "ann_fin_info": None, "third_party_flag": None, "third_party_list": None}  # 董事超额减持
annotations[164] = {"ann_related": 0, "ann_year": None, "ann_fin_flag": None, "ann_fin_info": None, "third_party_flag": None, "third_party_list": None}  # 合作协议未核实
annotations[165] = {"ann_related": 0, "ann_year": None, "ann_fin_flag": None, "ann_fin_info": None, "third_party_flag": None, "third_party_list": None}  # 资产冻结披露
annotations[166] = {"ann_related": 0, "ann_year": None, "ann_fin_flag": None, "ann_fin_info": None, "third_party_flag": None, "third_party_list": None}  # 股东增持违规
annotations[167] = {"ann_related": 0, "ann_year": None, "ann_fin_flag": None, "ann_fin_info": None, "third_party_flag": None, "third_party_list": None}  # 短线交易
annotations[168] = {"ann_related": 0, "ann_year": None, "ann_fin_flag": None, "ann_fin_info": None, "third_party_flag": None, "third_party_list": None}  # 治理不规范
annotations[169] = {"ann_related": 0, "ann_year": None, "ann_fin_flag": None, "ann_fin_info": None, "third_party_flag": None, "third_party_list": None}  # 厦华电子交易方案
annotations[170] = {"ann_related": 0, "ann_year": None, "ann_fin_flag": None, "ann_fin_info": None, "third_party_flag": None, "third_party_list": None}  # 股权转让
annotations[171] = {"ann_related": 1, "ann_year": [2020], "ann_fin_flag": 1, "ann_fin_info": [{"year": 2020, "elements": ["利润"]}], "third_party_flag": 0, "third_party_list": None}  # 会计差错更正调整净利润
annotations[172] = {"ann_related": 0, "ann_year": None, "ann_fin_flag": None, "ann_fin_info": None, "third_party_flag": None, "third_party_list": None}  # 股份回购
annotations[173] = {"ann_related": 0, "ann_year": None, "ann_fin_flag": None, "ann_fin_info": None, "third_party_flag": None, "third_party_list": None}  # 股东减持
annotations[174] = {"ann_related": 1, "ann_year": [2021], "ann_fin_flag": 1, "ann_fin_info": [{"year": 2021, "elements": ["收入"]}], "third_party_flag": 0, "third_party_list": None}  # 总额法调整为净额法导致收入多确认
annotations[175] = {"ann_related": 1, "ann_year": [2011], "ann_fin_flag": 0, "ann_fin_info": None, "third_party_flag": 0, "third_party_list": None}  # 关联交易未及时披露
annotations[176] = {"ann_related": 0, "ann_year": None, "ann_fin_flag": None, "ann_fin_info": None, "third_party_flag": None, "third_party_list": None}  # 股权转让收益确认
annotations[177] = {"ann_related": 1, "ann_year": [2018, 2019, 2020], "ann_fin_flag": 1, "ann_fin_info": [{"year": 2018, "elements": ["资产"]}, {"year": 2019, "elements": ["资产"]}, {"year": 2020, "elements": ["资产"]}], "third_party_flag": 0, "third_party_list": None}  # 实际控制人资金占用
annotations[178] = {"ann_related": 0, "ann_year": None, "ann_fin_flag": None, "ann_fin_info": None, "third_party_flag": None, "third_party_list": None}  # 内控缺陷
annotations[179] = {"ann_related": 0, "ann_year": None, "ann_fin_flag": None, "ann_fin_info": None, "third_party_flag": None, "third_party_list": None}  # 违规担保
annotations[180] = {"ann_related": 0, "ann_year": None, "ann_fin_flag": None, "ann_fin_info": None, "third_party_flag": None, "third_party_list": None}  # 虾夷扇贝苗底播计划审批程序问题
annotations[181] = {"ann_related": 0, "ann_year": None, "ann_fin_flag": None, "ann_fin_info": None, "third_party_flag": None, "third_party_list": None}  # 内幕交易
annotations[182] = {"ann_related": 0, "ann_year": None, "ann_fin_flag": None, "ann_fin_info": None, "third_party_flag": None, "third_party_list": None}  # 增持承诺未履行
annotations[183] = {"ann_related": 1, "ann_year": [2023], "ann_fin_flag": 0, "ann_fin_info": None, "third_party_flag": 0, "third_party_list": None}  # 关联交易未审议披露
annotations[184] = {"ann_related": 0, "ann_year": None, "ann_fin_flag": None, "ann_fin_info": None, "third_party_flag": None, "third_party_list": None}  # 短线交易（工作人员误操作）
annotations[185] = {"ann_related": 1, "ann_year": [2022, 2023, 2024], "ann_fin_flag": 1, "ann_fin_info": [{"year": 2022, "elements": ["收入"]}, {"year": 2023, "elements": ["收入"]}, {"year": 2024, "elements": ["收入"]}], "third_party_flag": 0, "third_party_list": None}  # 总额法确认收入导致收入确认不准确

# Build output dataframe
out_df = df.copy()

# Helper to convert to JSON string or None
def to_json_or_none(val):
    if val is None:
        return None
    return json.dumps(val, ensure_ascii=False)

# Fill annotation columns
out_df['ann_related'] = [annotations.get(i, {}).get('ann_related', None) for i in range(len(df))]
out_df['ann_year'] = [to_json_or_none(annotations.get(i, {}).get('ann_year', None)) for i in range(len(df))]
out_df['ann_fin_flag'] = [annotations.get(i, {}).get('ann_fin_flag', None) for i in range(len(df))]
out_df['ann_fin_info'] = [to_json_or_none(annotations.get(i, {}).get('ann_fin_info', None)) for i in range(len(df))]
out_df['third_party_flag'] = [annotations.get(i, {}).get('third_party_flag', None) for i in range(len(df))]
out_df['third_party_list'] = [to_json_or_none(annotations.get(i, {}).get('third_party_list', None)) for i in range(len(df))]

# Save
out_df.to_excel('MiniMax/minimax_ver.xlsx', index=False)
print("Saved MiniMax/minimax_ver.xlsx")

# Stats
r_related = sum(1 for i in range(len(df)) if annotations.get(i, {}).get('ann_related') == 1)
r_fin = sum(1 for i in range(len(df)) if annotations.get(i, {}).get('ann_fin_flag') == 1)
r_tp = sum(1 for i in range(len(df)) if annotations.get(i, {}).get('third_party_flag') == 1)
print(f"ann_related=1: {r_related}, ann_fin_flag=1: {r_fin}, third_party_flag=1: {r_tp}")