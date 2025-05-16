# # 国内航班航司数据
# DOM_AIRLINES = {
#     '主基地': {
#         '川航集团': {'ARR': 228, 'DEP': 228},
#         '国航集团': {'ARR': 244, 'DEP': 244}
#     },
#     '非主基地': {
#         '东航集团': {'ARR': 117, 'DEP': 117},
#         '海航集团': {'ARR': 119, 'DEP': 119},
#         '南航集团': {'ARR': 89, 'DEP': 89},
#         '其它': {'ARR': 69, 'DEP': 69}
#     }
# }
#
# # 国际航班航司数据
# INT_AIRLINES = {
#     '主基地': {
#         '川航集团': {'ARR': 81, 'DEP': 81},
#         '国航集团': {'ARR': 43, 'DEP': 42}
#     },
#     '非主基地': {
#         '东航集团': {'ARR': 5, 'DEP': 5},
#         '海航集团': {'ARR': 5, 'DEP': 5},
#         '南航集团': {'ARR': 0, 'DEP': 0},
#         '其它': {'ARR': 44, 'DEP': 44}
#     }
# }
#
# # 进港航向数据
# HEADINGS_ARR = {
#     'AKOPI': {'DOM': 266, 'INT': 9, 'INT性质': '近程'},
#     'BUPMI': {'DOM': 227, 'INT': 18, 'INT性质': '非绝对远程'},
#     'ELDUD': {'DOM': 162, 'INT': 34, 'INT性质': '非绝对远程'},
#     'IGNAK': {'DOM': 53, 'INT': 89, 'INT性质': '非绝对远程'},
#     'LADUP': {'DOM': 30, 'INT': 5, 'INT性质': '近程'},
#     'MEXAD': {'DOM': 128, 'INT': 23, 'INT性质': '绝对远程'}
# }
#
# # 离港航向数据
# HEADINGS_DEP = {
#     'ATVAX': {'DOM': 165, 'INT': 21, 'INT性质': '近程'},
#     'BOKIR': {'DOM': 142, 'INT': 23, 'INT性质': '绝对远程'},
#     'LUVEN': {'DOM': 50, 'INT': 104, 'INT性质': '非绝对远程'},
#     'MUMGO': {'DOM': 36, 'INT': 2, 'INT性质': '近程'},
#     'SAGPI': {'DOM': 213, 'INT': 7, 'INT性质': '绝对远程'},
#     'UBRAB': {'DOM': 260, 'INT': 20, 'INT性质': '近程'}
# }
#
# # 航司宽体机数据
# AIRLINES_WIDE = {
#     '川航集团': {'DOM': 36, 'INT': 37},
#     '东航集团': {'DOM': 3, 'INT': 0},
#     '国航集团': {'DOM': 22, 'INT': 10},
#     '海航集团': {'DOM': 23, 'INT': 1},
#     '南航集团': {'DOM': 2, 'INT': 0},
#     '其它': {'DOM': 1, 'INT': 14},
# }
#
# # 绝对远程航向航司集团
# ABSOLUTE_LONG_ROUTING = ['东航集团', '南航集团', '海航集团']
#
# # 不能有宽体机的航向
# ARR_DOM_WIDE_EXCEPTION_ROUTING = ['IGNAK']
# ARR_INT_WIDE_EXCEPTION_ROUTING = ['LADUP']
# DEP_DOM_WIDE_EXCEPTION_ROUTING = ['MUMGO']
# DEP_INT_WIDE_EXCEPTION_ROUTING = ['LUVEN']
#
# # 宽体机比例升高的航向
# ARR_DOM_WIDE_UP_ROUTING = []
# ARR_INT_WIDE_UP_ROUTING = []
# DEP_DOM_WIDE_UP_ROUTING = ['LUVEN', 'SAGPI']
# DEP_INT_WIDE_UP_ROUTING = []
#
# # 主航向比例约束例外航司集团
# MAIN_HEADING_EXCEPTION_AIRLINES = ['海航集团', '其它']
#
# # 波形约束例外航司集团
# WAVE_EXCEPTION_AIRLINES = ['海航集团', '其它']