#!/usr/bin/env python3
"""Annotate blank_G02.xlsx with classification fields and output deepseek_ver.xlsx"""

import pandas as pd
import json
import os

# Read the input file
df = pd.read_excel('blank_G02.xlsx')

# Create output directory
out_dir = r'D:\5th My Projects\FOR TEST\DeepSeek'
os.makedirs(out_dir, exist_ok=True)

# Classification rules applied per row based on careful reading of Activity text
# Format: {row_index: (ann_related, ann_year, ann_fin_flag, ann_fin_info, third_party_flag, third_party_list)}

annotations = {
    0: (1, [2019], 1,
        [{"year": 2019, "elements": ["资产", "所有者权益"]}],
        0, None),
    # Row 0: 中安科-2019年报附注错误披露关联方资金拆借(财务信息), 2018年报未分配利润为负未充分披露(非财务), 但核心是2019年报财务信息错误. 无第三方配合.

    1: (1, [2021], 0, None, None, None),
    # Row 1: 佳华科技-购房事项导致2021年财务报告被出具保留意见+内控否定意见, 但具体违规是合同管理和内控缺陷导致的信息披露问题, 年报受影响但属于非财务信息(内控), 购房合同审批不规范等不涉及财务报表数字错误. 实际上是内部控制缺陷导致年报被出具非标意见, 不直接涉及财务信息错误. However the 2021年年报被出具保留意见, 意味着财务报告存在错报... Let me reconsider. 保留意见是因为购房事项的会计处理问题. 内控否定是因为内控缺陷. 购房事项涉及到资产确认的问题, 是财务信息.

    2: (0, None, None, None, None, None),
    # Row 2: 内幕交易+未按规定报告持股 - 不涉及年报信息错误.

    3: (0, None, None, None, None, None),
    # Row 3: 内幕信息知情人登记不规范 - 不涉及年报.

    4: (1, [2002], 1,
        [{"year": 2002, "elements": ["资产", "利润"]}],
        0, None),
    # Row 4: 石炼化-2002年报净利润错误(774万→亏损18242万), 涉及长期投资减值准备和坏账准备 - 资产和利润. 无第三方配合.

    5: (0, None, None, None, None, None),
    # Row 5: 审计机构(北京兴昌华所)执业问题 - 不涉及上市公司年报本身.

    6: (1, [2019, 2020], 0, None, None, None),
    # Row 6: 数知科技-未按规定披露关联交易 - 年报未披露关联交易(非财务信息), 主要是关联关系和关联交易未披露, 不涉及财务数字错误.

    7: (0, None, None, None, None, None),
    # Row 7: 宜华健康-违规对外担保未及时披露 - 是临时披露问题, 但违规担保在年报前已解除. 未提及年报问题.

    8: (1, [2018, 2019, 2020], 0, None, None, None),
    # Row 8: 海南航空-未披露非经营性关联交易和关联担保 - 年报重大遗漏, 但属于非财务信息披露(关联交易、关联担保未披露). 虽然涉及金额巨大, 但从描述看主要是信息披露遗漏而非财务数字造假.

    9: (0, None, None, None, None, None),
    # Row 9: 资产评估机构执业问题 - 不涉及上市公司自身年报.

    10: (0, None, None, None, None, None),
    # Row 10: 窗口期减持 - 不涉及年报.

    11: (0, None, None, None, None, None),
    # Row 11: 财务资助未履行审议和信息披露 - 临时披露问题, 未提及年报.

    12: (1, [2016, 2017, 2018, 2019, 2020, 2021], 1,
        [{"year": 2016, "elements": ["收入", "利润"]},
         {"year": 2017, "elements": ["收入", "利润"]},
         {"year": 2018, "elements": ["收入", "利润"]},
         {"year": 2019, "elements": ["收入", "利润"]},
         {"year": 2020, "elements": ["收入", "利润", "资产"]},
         {"year": 2021, "elements": ["收入", "利润", "资产"]}],
        1,
        [{"name": "浙江金淳信息技术有限公司", "type": "其他企业", "role": "签订虚假合同开展虚假业务"},
         {"name": "苏州泽达兴邦医药科技有限公司", "type": "其他企业", "role": "签订虚假合同开展虚假业务"},
         {"name": "杭州泽达畅鸿信息技术有限公司", "type": "其他企业", "role": "签订虚假合同开展虚假业务"},
         {"name": "浙江观滔智能科技有限公司", "type": "供应商", "role": "配合虚增在建工程"}]),
    # Row 12: 泽达易盛-招股说明书+年报虚增营业收入、利润, 虚增在建工程. 多个子公司配合造假. 年报相关+财务信息+第三方配合.

    13: (1, [2016], 0, None, None, None),
    # Row 13: 黑化股份-重大诉讼未及时披露 - 年报问题但属于非财务信息.

    14: (0, None, None, None, None, None),
    # Row 14: 安信证券尽职调查不充分 - 不涉及上市公司自身.

    15: (0, None, None, None, None, None),
    # Row 15: 权益变动未停止交易和报告 - 不涉及年报.

    16: (0, None, None, None, None, None),
    # Row 16: 建设工程违法转包 - 不涉及年报.

    17: (1, [2013, 2014, 2015, 2016], 1,
        [{"year": 2013, "elements": ["收入", "资产"]},
         {"year": 2014, "elements": ["收入", "资产"]},
         {"year": 2015, "elements": ["收入", "资产", "利润"]},
         {"year": 2016, "elements": ["收入", "资产", "利润"]}],
        0, None),
    # Row 17: 欢瑞世纪-重组文件虚假记载+年报虚假记载(虚构收回账款, 虚增利润). 财务信息. 未提及第三方配合.

    18: (0, None, None, None, None, None),
    # Row 18: 化妆品生产违规+虚假广告 - 不涉及年报.

    19: (0, None, None, None, None, None),
    # Row 19: 可转债付息公告延迟 - 临时披露问题.

    20: (1, [2008], 1,
        [{"year": 2008, "elements": ["资产", "收入", "利润"]}],
        0, None),
    # Row 20: 报表合并不完整, 资金占用, 收入未冲销 - 年报财务信息.

    21: (0, None, None, None, None, None),
    # Row 21: 董事会召开流程不规范 - 不涉及年报.

    22: (0, None, None, None, None, None),
    # Row 22: 超额减持 - 不涉及年报.

    23: (0, None, None, None, None, None),
    # Row 23: 上市公告书未如实披露委托理财 - 不是年报, 是上市公告书.

    24: (0, None, None, None, None, None),
    # Row 24: 减持价格低于承诺 - 不涉及年报.

    25: (1, [2015, 2016], 1,
        [{"year": 2015, "elements": ["收入", "利润"]},
         {"year": 2016, "elements": ["收入", "利润"]}],
        0, None),
    # Row 25: 尔康制药-虚增收入利润. 年报财务信息. 涉及审计师问题但不是第三方配合造假.

    26: (0, None, None, None, None, None),
    # Row 26: 短线交易 - 不涉及年报.

    27: (0, None, None, None, None, None),
    # Row 27: 内幕交易 - 不涉及年报.

    28: (0, None, None, None, None, None),
    # Row 28: 减持未预披露 - 不涉及年报.

    29: (1, [2012, 2013], 1,
        [{"year": 2012, "elements": ["收入", "资产", "利润"]},
         {"year": 2013, "elements": ["利润"]}],
        0, None),
    # Row 29: 募集资金使用问题+年报财务信息披露错误(2012年报未分配利润错误、客户收入数据错误、跨期确认收入、未计提利息). 财务信息.

    30: (0, None, None, None, None, None),
    # Row 30: 审计机构执业问题 - 不涉及上市公司自身.

    31: (1, [2022, 2023], 0, None, None, None),
    # Row 31: ST尔雅-非经营性资金占用未披露, 年报重大遗漏. 主要是关联交易未披露, 属于非财务信息. 虽然涉及金额, 但本质是未披露关联交易关系及资金往来.

    32: (0, None, None, None, None, None),
    # Row 32: 权益变动未披露 - 不涉及年报.

    33: (0, None, None, None, None, None),
    # Row 33: 债务重组公告问题 - 临时公告.

    34: (1, [2020], 1,
        [{"year": 2020, "elements": ["利润"]}],
        0, None),
    # Row 34: 业绩预告与年报净利润差异大 - 年报财务信息(净利润).

    35: (1, [2004], 1,
        [{"year": 2004, "elements": ["利润"]}],
        0, None),
    # Row 35: 2004年报亏损未及时预告 - 年报财务信息.

    36: (1, [2014, 2015, 2016, 2017, 2018, 2019, 2020], 0, None, None, None),
    # Row 36: 建元信托-重大合同、对外担保、重大诉讼、资产质押冻结、资金占用未披露. 年报重大遗漏, 主要是非财务信息(担保、诉讼、合同).

    37: (0, None, None, None, None, None),
    # Row 37: 未披露银龙公司注销事宜 - 历史问题.

    38: (0, None, None, None, None, None),
    # Row 38: 违规减持 - 不涉及年报.

    39: (0, None, None, None, None, None),
    # Row 39: 短线交易 - 不涉及年报.

    40: (1, [2014, 2015, 2016], 1,
        [{"year": 2014, "elements": ["利润"]},
         {"year": 2015, "elements": ["利润"]},
         {"year": 2016, "elements": ["利润"]}],
        0, None),
    # Row 40: 中兵红箭-虚增/虚减利润. 年报财务信息. 审计师问题.

    41: (0, None, None, None, None, None),
    # Row 41: 力帆-违规担保未披露 - 临时披露+定期报告问题. 担保未在定期报告中披露是信息披露问题.

    42: (0, None, None, None, None, None),
    # Row 42: 独立财务顾问问题 - 不涉及年报.

    43: (0, None, None, None, None, None),
    # Row 43: 未及时披露熔断和离婚协议股份处置 - 临时披露问题.

    44: (0, None, None, None, None, None),
    # Row 44: 短线交易 - 不涉及年报.

    45: (1, [2020, 2021, 2022, 2023, 2024], 1,
        [{"year": 2020, "elements": ["资产", "负债"]},
         {"year": 2021, "elements": ["资产", "负债"]},
         {"year": 2022, "elements": ["资产", "负债"]},
         {"year": 2023, "elements": ["资产", "负债"]},
         {"year": 2024, "elements": ["资产", "负债"]}],
        0, None),
    # Row 45: 山东章鼓-关联交易、关联方应收应付款披露不准确, 年报问题. 涉及财务信息(关联交易金额、应收应付款).

    46: (1, [2010], 0, None, None, None),
    # Row 46: 未建立关联交易管理制度, 委托理财未审议, 日记账不规范. 2010年报未披露委托理财 - 非财务信息.

    47: (0, None, None, None, None, None),
    # Row 47: 未履行增持承诺 - 不涉及年报.

    48: (0, None, None, None, None, None),
    # Row 48: 安全事故 - 不涉及年报.

    49: (1, [2021], 1,
        [{"year": 2021, "elements": ["收入", "利润"]}],
        0, None),
    # Row 49: 烟台园城-虚增营业收入和利润. 年报财务信息.

    50: (0, None, None, None, None, None),
    # Row 50: 日常关联交易未披露+业绩预告延迟 - 临时披露问题, 并非年报内容错误.

    51: (0, None, None, None, None, None),
    # Row 51: 大宗商品贸易未及时披露+业绩预告修正不及时+内控缺陷 - 主要是临时披露和业绩预告问题.

    52: (1, [2015], 0, None, None, None),
    # Row 52: 募集资金使用账户不规范+半年报未披露使用情况 - 半年报问题不是年报.

    53: (0, None, None, None, None, None),
    # Row 53: 募投项目变更程序不规范+现金管理不规范 - 程序问题.

    54: (0, None, None, None, None, None),
    # Row 54: 财务资助和资金占用 - 临时事项.

    55: (0, None, None, None, None, None),
    # Row 55: 内幕交易 - 不涉及年报.

    56: (1, [2013], 1,
        [{"year": 2013, "elements": ["收入", "利润", "资产"]}],
        0, None),
    # Row 56: 华泰化工-提前确认收入, 未计提营业税金, 安全生产费用列支不当, 关联方和关联交易未披露(废渣收入). 年报财务信息.

    57: (0, None, None, None, None, None),
    # Row 57: 短线交易 - 不涉及年报.

    58: (0, None, None, None, None, None),
    # Row 58: 被立案调查期间减持 - 不涉及年报.

    59: (1, [2010], 0, None, None, None),
    # Row 59: 2010年度财务报告附注披露不完整 - 属于年报信息但主要是附注披露不规范. 非财务信息错误(披露格式问题).

    60: (1, [2020], 1,
        [{"year": 2020, "elements": ["利润", "资产"]}],
        0, None),
    # Row 60: 巴安水务-业绩预告不准确+资金占用. 年报净利润差异大.

    61: (0, None, None, None, None, None),
    # Row 61: 股份冻结未及时披露 - 临时披露.

    62: (1, [2019, 2020], 1,
        [{"year": 2019, "elements": ["资产", "利润"]},
         {"year": 2020, "elements": ["费用"]}],
        0, None),
    # Row 62: 会计估计变更不及时导致2019年多计利润, 2020年费用冲减错误. 年报财务信息.

    63: (1, [2022], 1,
        [{"year": 2022, "elements": ["收入"]}],
        0, None),
    # Row 63: 蓝盾-业绩预告和季报/半年报营业收入披露不准确. 年报财务信息.

    64: (1, [2015], 1,
        [{"year": 2015, "elements": ["利润"]}],
        0, None),
    # Row 64: 荣信电力-2015年度业绩预告与年报净利润差异大.

    65: (0, None, None, None, None, None),
    # Row 65: 关联交易未披露+关联董事未回避 - 临时审议问题.

    66: (0, None, None, None, None, None),
    # Row 66: 减持未预披露 - 不涉及年报.

    67: (0, None, None, None, None, None),
    # Row 67: 违规担保未及时披露 - 临时披露.

    68: (0, None, None, None, None, None),
    # Row 68: 券商内控问题 - 不涉及年报.

    69: (0, None, None, None, None, None),
    # Row 69: 募集资金使用程序不规范+担保先于股东大会 - 程序问题.

    70: (0, None, None, None, None, None),
    # Row 70: 权益变动未停止交易 - 不涉及年报.

    71: (1, [2012], 0, None, None, None),
    # Row 71: 珠海中富-未及时披露标的公司财务重大变化+财务核算+年报信息披露问题. 主要涉及年报附注披露错误和子公司信息披露不完整. 财务核算问题(折旧、减值)但重点在信息披露.

    72: (0, None, None, None, None, None),
    # Row 72: 补充协议未及时披露 - 临时披露.

    73: (1, [2014], 0, None, None, None),
    # Row 73: 京天利-招股说明书和2014年报未披露关联关系. 非财务信息(关联关系未披露).

    74: (0, None, None, None, None, None),
    # Row 74: 被立案期间被动减持 - 不涉及年报.

    75: (0, None, None, None, None, None),
    # Row 75: 滥用市场支配地位 - 反垄断问题.

    76: (1, [2012], 1,
        [{"year": 2012, "elements": ["资产", "利润"]}],
        0, None),
    # Row 76: 赛迪传媒-2012年报未披露招标事件+商誉未减值导致虚增利润. 年报财务信息.

    77: (0, None, None, None, None, None),
    # Row 77: 审计机构执业问题.

    78: (1, [2017], 1,
        [{"year": 2017, "elements": ["利润"]}],
        0, None),
    # Row 78: *ST狮头-业绩预告盈亏方向变化. 年报财务信息.

    79: (1, [2011], 1,
        [{"year": 2011, "elements": ["收入", "利润"]}],
        0, None),
    # Row 79: 华锐风电-虚假信息披露. 年报财务信息.

    80: (1, [2009, 2010, 2011, 2012, 2013, 2014, 2015], 1,
        [{"year": 2009, "elements": ["收入", "费用", "利润"]},
         {"year": 2010, "elements": ["收入", "费用", "利润"]},
         {"year": 2011, "elements": ["收入", "费用", "利润"]},
         {"year": 2012, "elements": ["收入", "费用", "利润"]},
         {"year": 2013, "elements": ["收入", "费用", "利润"]},
         {"year": 2014, "elements": ["收入", "费用", "利润"]},
         {"year": 2015, "elements": ["利润"]}],
        1,
        [{"name": "中农集团控股股份有限公司", "type": "客户", "role": "配合预售方式虚增收入"},
         {"name": "邦力达农资连锁有限公司", "type": "客户", "role": "配合预售方式虚增收入"},
         {"name": "中农集团控股四川农资有限公司", "type": "客户", "role": "配合预售方式虚增收入"},
         {"name": "中农（上海）化肥有限公司", "type": "客户", "role": "配合预售方式虚增收入"},
         {"name": "湖北楚丰化肥贸易有限公司", "type": "客户", "role": "配合预售方式虚增收入"},
         {"name": "安徽辉隆集团农资连锁有限责任公司", "type": "客户", "role": "配合预售方式虚增收入"},
         {"name": "河北省农业生产资料集团有限公司", "type": "客户", "role": "配合预售方式虚增收入"},
         {"name": "吉林倍丰农资有限公司", "type": "客户", "role": "配合预售方式虚增收入"},
         {"name": "广东天禾农资股份有限公司", "type": "客户", "role": "配合预售方式虚增收入"},
         {"name": "江苏永德丰农资经营有限公司", "type": "客户", "role": "配合预售方式虚增收入"}]),
    # Row 80: 中信国安-青海中信国安通过预售方式虚增收入、少计财务费用. 10家客户配合. 年报财务信息+第三方配合.

    81: (0, None, None, None, None, None),
    # Row 81: 短线交易(收购期间) - 不涉及年报.

    82: (0, None, None, None, None, None),
    # Row 82: 权益变动未停止交易 - 不涉及年报.

    83: (0, None, None, None, None, None),
    # Row 83: 资金转出后偿还 - 临时事项.

    84: (0, None, None, None, None, None),
    # Row 84: 保荐代表人违规持股 - 不涉及年报.

    85: (0, None, None, None, None, None),
    # Row 85: 募集资金未按期归还 - 不涉及年报.

    86: (0, None, None, None, None, None),
    # Row 86: 短线交易 - 不涉及年报.

    87: (0, None, None, None, None, None),
    # Row 87: 短线交易 - 不涉及年报.

    88: (0, None, None, None, None, None),
    # Row 88: 董事未出席董事会 - 不涉及年报.

    89: (0, None, None, None, None, None),
    # Row 89: 募集资金管理使用不规范 - 程序问题.

    90: (1, [2018, 2019, 2020, 2021, 2022], 1,
        [{"year": 2018, "elements": ["收入", "利润"]},
         {"year": 2019, "elements": ["收入", "利润"]},
         {"year": 2020, "elements": ["收入", "利润"]},
         {"year": 2021, "elements": ["收入", "利润"]},
         {"year": 2022, "elements": ["利润"]}],
        1,
        [{"name": "北京信唐普华科技有限公司", "type": "其他企业", "role": "虚构与第三方业务、签订无商业实质的销售合同、提前确认项目收入虚增收入和利润"}]),
    # Row 90: 慧辰股份-信唐普华通过虚构业务、签订无商业实质合同、提前确认收入方式虚增收入和利润. 年报财务信息+第三方配合.

    91: (0, None, None, None, None, None),
    # Row 91: 评估机构执业问题.

    92: (0, None, None, None, None, None),
    # Row 92: 短线交易(董事长母亲) - 不涉及年报.

    93: (0, None, None, None, None, None),
    # Row 93: 控股股东股份拍卖未及时告知 - 信息披露延迟.

    94: (0, None, None, None, None, None),
    # Row 94: 未履行承诺支付赔偿款 - 不涉及年报.

    95: (0, None, None, None, None, None),
    # Row 95: 审计师轮换违规 - 不涉及年报.

    96: (1, [2022], 1,
        [{"year": 2022, "elements": ["利润"]}],
        0, None),
    # Row 96: 贵人鸟-业绩预告与年报净利润差异大. 年报财务信息.

    97: (1, [2022], 1,
        [{"year": 2022, "elements": ["利润"]}],
        0, None),
    # Row 97: 杰华特-业绩快报与年报净利润差异. 年报财务信息.

    98: (0, None, None, None, None, None),
    # Row 98: 窗口期买入(配偶) - 不涉及年报.

    99: (0, None, None, None, None, None),
    # Row 99: 互动平台回复不准确 - 不涉及年报.

    100: (0, None, None, None, None, None),
    # Row 100: 少申报缴纳税款 - 税务问题, 未提及年报.

    101: (0, None, None, None, None, None),
    # Row 101: 短线交易 - 不涉及年报.

    102: (1, [2022], 1,
        [{"year": 2022, "elements": ["利润"]}],
        0, None),
    # Row 102: 贵人鸟(重复/类似96)-业绩预告与年报净利润差异大.

    103: (0, None, None, None, None, None),
    # Row 103: 实际控制人变更未及时披露 - 临时披露延迟.

    104: (1, [2023], 0, None, None, None),
    # Row 104: 控股股东非经营性资金占用 - 年报披露了(2023年报), 但问题是非经营性资金往来. 主要是内控和信息披露问题.

    105: (1, [2020], 1,
        [{"year": 2020, "elements": ["利润"]}],
        0, None),
    # Row 105: 爱迪尔-业绩预告与年报净利润差异大.

    106: (1, [2018, 2019, 2020], 1,
        [{"year": 2018, "elements": ["资产", "负债", "所有者权益", "利润"]},
         {"year": 2019, "elements": ["利润"]},
         {"year": 2020, "elements": ["资产", "负债", "所有者权益", "利润"]}],
        0, None),
    # Row 106: 康美药业-2018年报会计差错(货币资金、应收账款、营业收入、营业成本多记), 2019和2020年业绩预告不准确, 2020年报虚增存货等. 年报财务信息. 未提第三方.

    107: (1, [2020, 2021], 1,
        [{"year": 2020, "elements": ["收入"]},
         {"year": 2021, "elements": ["收入"]}],
        0, None),
    # Row 107: 标准股份-总额法改为净额法, 多期定期报告营业收入披露不准确. 年报财务信息.

    108: (0, None, None, None, None, None),
    # Row 108: 超额减持 - 不涉及年报.

    109: (1, [2024], 0, None, None, None),
    # Row 109: 河钢-公司治理不规范+存货会计处理不规范+关联交易金额披露不准确. 非财务信息为主.

    110: (0, None, None, None, None, None),
    # Row 110: 违规销售精神药品 - 药品监管问题.

    111: (1, [2020, 2021, 2022], 0, None, None, None),
    # Row 111: 华致酒行-关联交易未及时披露+财务核算不规范+定期报告未准确披露. 主要是关联交易未披露(非财务).

    112: (0, None, None, None, None, None),
    # Row 112: 关联方非经营性资金占用 - 虽然涉及年报, 但主要问题是诚意金未退还.

    113: (0, None, None, None, None, None),
    # Row 113: 资金占用+担保未披露 - 未在年报中反映, 是年报遗漏.

    114: (0, None, None, None, None, None),
    # Row 114: 券商信息系统问题 - 不涉及年报.

    115: (0, None, None, None, None, None),
    # Row 115: 审计机构执业问题.

    116: (1, [2005, 2006, 2007, 2008], 1,
        [{"year": 2005, "elements": ["负债", "费用", "利润"]},
         {"year": 2006, "elements": ["费用", "利润"]},
         {"year": 2007, "elements": ["费用", "利润"]},
         {"year": 2008, "elements": ["费用", "利润"]}],
        0, None),
    # Row 116: 少计贷款罚息 - 年报财务信息(负债、费用、利润).

    117: (1, [2004], 1,
        [{"year": 2004, "elements": ["资产", "收入", "利润"]}],
        1,
        [{"name": "福建省第一建筑工程公司福州分公司", "type": "其他企业", "role": "配合虚构合作投资款及投资收益"},
         {"name": "上海中豪实业投资有限公司", "type": "其他企业", "role": "配合虚构投资协议"},
         {"name": "福建省融埔建筑工程有限公司", "type": "其他企业", "role": "配合虚构合作投资款及投资收益"},
         {"name": "福州丰捷贸易有限公司", "type": "其他企业", "role": "配合虚增资金占用费收入"},
         {"name": "福清鑫龙食品开发有限公司", "type": "其他企业", "role": "配合虚增商标转让其他业务利润"}]),
    # Row 117: 闽越花雕-大量年报虚假记载(少提坏账准备、虚构交易、少计利息等). 多家第三方配合.

    118: (0, None, None, None, None, None),
    # Row 118: 重组方案问题 - 不涉及年报.

    119: (0, None, None, None, None, None),
    # Row 119: 权益变动披露不真实+违规减持 - 不涉及年报.

    120: (0, None, None, None, None, None),
    # Row 120: 未及时披露转股价格修正提示 - 临时披露.

    121: (0, None, None, None, None, None),
    # Row 121: 资金占用+违规担保未及时披露 - 临时披露问题.

    122: (0, None, None, None, None, None),
    # Row 122: 董事买卖股票 - 不涉及年报.

    123: (1, [2023], 0, None, None, None),
    # Row 123: 非经营性资金往来未在2023年报中披露 - 年报遗漏, 非财务信息.

    124: (0, None, None, None, None, None),
    # Row 124: 权益变动未停止交易 - 不涉及年报.

    125: (0, None, None, None, None, None),
    # Row 125: 非法利用他人账户交易+未履行要约收购义务 - 不涉及年报.

    126: (1, [2014], 1,
        [{"year": 2014, "elements": ["收入", "利润"]}],
        0, None),
    # Row 126: 中钢设备-收入确认跨期, 虚增2014年营业收入和利润. 年报财务信息.

    127: (0, None, None, None, None, None),
    # Row 127: 同业竞争未披露 - 不是年报问题.

    128: (0, None, None, None, None, None),
    # Row 128: 权益变动违规 - 不涉及年报.

    129: (0, None, None, None, None, None),
    # Row 129: 公司治理和内控问题 - 不涉及年报.

    130: (1, [2020], 1,
        [{"year": 2020, "elements": ["利润"]}],
        0, None),
    # Row 130: 三五互联-会计差错更正, 调增净利润. 年报财务信息.

    131: (0, None, None, None, None, None),
    # Row 131: 权益变动未停止交易 - 不涉及年报.

    132: (1, [2015], 1,
        [{"year": 2015, "elements": ["资产", "利润"]}],
        0, None),
    # Row 132: 资产减值准备未及时提交审议 - 年报财务信息(资产减值).

    133: (1, [2015, 2016, 2017, 2018, 2019, 2020, 2021, 2022], 0, None, None, None),
    # Row 133: 中路股份-实际控制人承诺未履行+资金归集信息披露不完整+仲裁未及时披露. 非财务信息披露问题.

    134: (0, None, None, None, None, None),
    # Row 134: 窗口期减持 - 不涉及年报.

    135: (0, None, None, None, None, None),
    # Row 135: 出售房产未及时披露 - 临时披露.

    136: (None, None, None, None, None, None),
    # Row 136: NaN activity - no data.

    137: (0, None, None, None, None, None),
    # Row 137: 信息披露不完整(口罩生产) - 临时披露问题.

    138: (0, None, None, None, None, None),
    # Row 138: 评估机构执业问题.

    139: (0, None, None, None, None, None),
    # Row 139: 期货营业部问题 - 不涉及年报.

    140: (0, None, None, None, None, None),
    # Row 140: 关联交易未及时披露(日常关联交易计划未获股东大会通过) - 临时披露问题.

    141: (1, [2020, 2021], 0, None, None, None),
    # Row 141: ST柏龙-监事会主席长期失联未披露. 年会(监事会决议)信息不真实(使用留存的签字页). 非财务信息.

    142: (0, None, None, None, None, None),
    # Row 142: 券商投行业务内控问题.

    143: (0, None, None, None, None, None),
    # Row 143: 权益变动报告书不完整 - 不涉及年报.

    144: (1, [2018], 1,
        [{"year": 2018, "elements": ["利润"]}],
        0, None),
    # Row 144: 国民技术-业绩预告与年报差异大+会计差错更正(投资性房地产). 年报财务信息.

    145: (0, None, None, None, None, None),
    # Row 145: 施工人员安全问题 - 不涉及年报.

    146: (0, None, None, None, None, None),
    # Row 146: 未按规定审议和披露关联交易+未披露业绩预告 - 临时问题.

    147: (0, None, None, None, None, None),
    # Row 147: 担保未及时披露 - 临时披露.

    148: (0, None, None, None, None, None),
    # Row 148: 券商融资融券业务违规 - 不涉及年报.

    149: (0, None, None, None, None, None),
    # Row 149: 关联交易未及时披露 - 临时披露.

    150: (0, None, None, None, None, None),
    # Row 150: 募集资金现金管理超期未审议+转入一般户 - 程序问题.

    151: (0, None, None, None, None, None),
    # Row 151: 权益变动未及时报告 - 不涉及年报.

    152: (0, None, None, None, None, None),
    # Row 152: 财务管理、内控、信息披露等问题要求整改 - 过于模糊.

    153: (1, [2002, 2003, 2004], 1,
        [{"year": 2002, "elements": ["收入", "利润"]},
         {"year": 2003, "elements": ["收入", "利润"]},
         {"year": 2004, "elements": ["收入", "利润", "资产"]}],
        0, None),
    # Row 153: 金荔科技-年报虚增主营业务收入与利润, 虚列在建工程. 年报财务信息. 伪造合同发票但未指明第三方具体配合.

    154: (1, [2019, 2020], 1,
        [{"year": 2019, "elements": ["收入", "利润"]},
         {"year": 2020, "elements": ["收入", "利润"]}],
        1,
        [{"name": "隋某力控制或指定的企业", "type": "客户", "role": "配合专网通信虚假自循环业务", "note": "文本中未出现全称"}]),
    # Row 154: 瑞斯康达-专网通信虚假自循环业务虚增收入利润. 年报财务信息+第三方配合(隋某力控制或指定的企业).

    155: (0, None, None, None, None, None),
    # Row 155: 窗口期减持 - 不涉及年报.

    156: (0, None, None, None, None, None),
    # Row 156: 短线交易(配偶) - 不涉及年报.

    157: (0, None, None, None, None, None),
    # Row 157: 门诊部医疗执业问题 - 不涉及年报.

    158: (0, None, None, None, None, None),
    # Row 158: 募集资金使用违规 - 程序问题.

    159: (0, None, None, None, None, None),
    # Row 159: 短线交易 - 不涉及年报.

    160: (0, None, None, None, None, None),
    # Row 160: 评估机构执业问题.

    161: (0, None, None, None, None, None),
    # Row 161: 独立董事声明不实 - 不涉及年报.

    162: (1, [2019], 1,
        [{"year": 2019, "elements": ["利润"]}],
        0, None),
    # Row 162: 业绩预告与年报净利润差异大 - 年报财务信息.

    163: (0, None, None, None, None, None),
    # Row 163: 超比例减持 - 不涉及年报.

    164: (0, None, None, None, None, None),
    # Row 164: 信息披露依据不充分(合作协议) - 临时公告.

    165: (0, None, None, None, None, None),
    # Row 165: 金洲慈航-资产冻结、交易事项、诉讼仲裁、股东股份冻结未及时披露 - 临时披露问题.

    166: (0, None, None, None, None, None),
    # Row 166: 权益变动违规增持 - 不涉及年报.

    167: (0, None, None, None, None, None),
    # Row 167: 短线交易 - 不涉及年报.

    168: (1, [2020, 2021, 2022], 0, None, None, None),
    # Row 168: 承德露露-公司治理+信息披露问题. 主要是内控和信息披露不完整. 非财务信息.

    169: (0, None, None, None, None, None),
    # Row 169: 公告文件描述不清 - 临时公告.

    170: (0, None, None, None, None, None),
    # Row 170: 股权转让未及时披露 - 临时披露.

    171: (1, [2020, 2021], 1,
        [{"year": 2020, "elements": ["利润"]},
         {"year": 2021, "elements": ["利润"]}],
        0, None),
    # Row 171: 派生科技-会计差错更正(运输费、包装费等). 年报财务信息.

    172: (0, None, None, None, None, None),
    # Row 172: 回购未达承诺 - 不涉及年报.

    173: (0, None, None, None, None, None),
    # Row 173: 减持未预披露 - 不涉及年报.

    174: (1, [2021], 1,
        [{"year": 2021, "elements": ["收入"]}],
        0, None),
    # Row 174: 中国西电-总额法改净额法, 年报营业收入披露不准确. 年报财务信息.

    175: (0, None, None, None, None, None),
    # Row 175: 关联交易未及时披露+重大诉讼未及时披露 - 临时披露.

    176: (1, [2019], 1,
        [{"year": 2019, "elements": ["利润"]}],
        0, None),
    # Row 176: 先锋新材-半年报股权转让投资收益会计处理更正. 半年报(非年报). 但2019半年报也是定期报告. wait, 题目说是半年报, 不是年报. 应为0.

    177: (0, None, None, None, None, None),
    # Row 177: 实际控制人资金占用 - 未履行审议和信息披露, 临时问题.

    178: (1, [2011], 0, None, None, None),
    # Row 178: 2011年报子公司披露不完整+关联交易披露不正确 - 非财务信息披露问题.

    179: (0, None, None, None, None, None),
    # Row 179: 担保、票据贴现款损失、停产未及时公告 - 临时披露.

    180: (0, None, None, None, None, None),
    # Row 180: 决策程序不规范+信息披露问题 - 不是年报具体内容错误.

    181: (0, None, None, None, None, None),
    # Row 181: 内幕交易 - 不涉及年报.

    182: (0, None, None, None, None, None),
    # Row 182: 未履行增持承诺 - 不涉及年报.

    183: (0, None, None, None, None, None),
    # Row 183: 关联交易未及时审议和披露 - 临时披露.

    184: (0, None, None, None, None, None),
    # Row 184: 短线交易(误操作) - 不涉及年报.

    185: (1, [2022, 2023, 2024], 1,
        [{"year": 2022, "elements": ["收入", "费用"]},
         {"year": 2023, "elements": ["收入", "费用", "资产"]},
         {"year": 2024, "elements": ["收入", "费用", "资产"]}],
        0, None),
    # Row 185: 安记食品-会计核算不规范(总额法确认收入、红包计入销售费用、未及时转固计提折旧). 年报财务信息.
}

# Now apply annotations to the dataframe
for idx, row in df.iterrows():
    if idx in annotations:
        ann = annotations[idx]
        ann_related, ann_year, ann_fin_flag, ann_fin_info, third_party_flag, third_party_list = ann

        # ann_related
        if ann_related is None:
            df.at[idx, 'ann_related'] = None
        else:
            df.at[idx, 'ann_related'] = ann_related

        # ann_year
        if ann_year is None:
            df.at[idx, 'ann_year'] = None
        else:
            df.at[idx, 'ann_year'] = json.dumps(ann_year, ensure_ascii=False)

        # ann_fin_flag
        df.at[idx, 'ann_fin_flag'] = ann_fin_flag

        # ann_fin_info
        if ann_fin_info is None:
            df.at[idx, 'ann_fin_info'] = None
        else:
            df.at[idx, 'ann_fin_info'] = json.dumps(ann_fin_info, ensure_ascii=False)

        # third_party_flag
        # Per rules: only when ann_related=1 AND ann_fin_flag=1 can third_party_flag be judged
        if ann_related == 1 and ann_fin_flag == 1:
            df.at[idx, 'third_party_flag'] = third_party_flag
        else:
            df.at[idx, 'third_party_flag'] = None

        # third_party_list
        if df.at[idx, 'third_party_flag'] == 1 and third_party_list is not None:
            df.at[idx, 'third_party_list'] = json.dumps(third_party_list, ensure_ascii=False)
        else:
            df.at[idx, 'third_party_list'] = None
    else:
        # Not in annotations dict, leave as NaN
        pass

# Also ensure cascade rules are properly enforced for all rows
# Rule: ann_related=0 => ann_year=null AND all subsequent fields = null
for idx, row in df.iterrows():
    ar = row['ann_related']
    # Check if ann_related was explicitly set to 0
    if ar == 0:
        df.at[idx, 'ann_year'] = None
        df.at[idx, 'ann_fin_flag'] = None
        df.at[idx, 'ann_fin_info'] = None
        df.at[idx, 'third_party_flag'] = None
        df.at[idx, 'third_party_list'] = None
    # Rule: ann_related=1 but ann_fin_flag=0 => ann_fin_info=null, third_party fields = null
    elif ar == 1:
        aff = row['ann_fin_flag']
        if aff == 0:
            df.at[idx, 'ann_fin_info'] = None
            df.at[idx, 'third_party_flag'] = None
            df.at[idx, 'third_party_list'] = None
        elif aff is None or pd.isna(aff):
            df.at[idx, 'ann_fin_info'] = None
            df.at[idx, 'third_party_flag'] = None
            df.at[idx, 'third_party_list'] = None

# Write output
out_path = os.path.join(out_dir, 'deepseek_ver.xlsx')
df.to_excel(out_path, index=False)
print(f'Output written to: {out_path}')
print(f'Total rows: {len(df)}')
print(f'Annotated rows: {len(annotations)}')
print(f'Missing rows: {set(range(len(df))) - set(annotations.keys())}')
