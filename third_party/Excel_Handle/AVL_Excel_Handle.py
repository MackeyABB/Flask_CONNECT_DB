'''
Introduction:
    此程序用于AVL Excel文件的相关处理操作。
Usage:

Notes:

Revision log:
1.0.0 - 202601117: 初始版本，实现基本功能。
'''


# 版本号
# xx.yy.zz
# xx: 大版本，架构性变化
# yy: 功能性新增
# zz: Bug修复
__revision__ = '1.0.0'


import openpyxl

# Excel列与sql_result索引的对应关系
# sql result索引关系见third_party\PyPika_CONNECT\PyPika\PyPika_CONNECT.py中的FIELDS_SAPMaxDB， FIELDS_AccessDB
# 第一个值表示Excel的列，第二个值表示sql_result的索引，从0开始计数，None表示该列为空或不处理
excel_mapping = [
    ('A', None),        # 序号，特殊处理
    ('B', 2),           # SAP_Number
    ('C', 3),           # SAP_Description
    ('D', None),        # 空
    ('E', 6),           # manufact_1
    ('F', 7),           # manufact_partnum_1
    ('G', 9),           # manufact_2
    ('H', 10),          # manufact_partnum_2
    ('I', 12),          # manufact_3
    ('J', 13),          # manufact_partnum_3
    ('K', 15),          # manufact_4
    ('L', 16),          # manufact_partnum_4
    ('M', 18),          # manufact_5
    ('N', 19),          # manufact_partnum_5
    ('O', 21),          # manufact_6
    ('P', 22),          # manufact_partnum_6
    ('Q', 24),          # manufact_7
    ('R', 25),          # manufact_partnum_7
]

def first_write_AVL_to_excel(template_file, sql_result, output_excel_file):
    # 加载模板文件
    wb = openpyxl.load_workbook(template_file)
    # AVL sheet保存数据
    sheet_avl = wb['AVL']
    # AVL sheet保存数据
    for row_idx, row_data in enumerate(sql_result, start=7):  # Excel第7行开始
        for col_idx, (col_letter, data_idx) in enumerate(excel_mapping, start=1):
            cell = f"{col_letter}{row_idx}"
            if data_idx is None:
                # 特殊处理：A列为序号，D列为空
                if col_letter == 'A':
                    value = row_idx - 6
                else:
                    value = ""
            else:
                value = row_data[data_idx]
            sheet_avl[cell] = value
    wb.save(output_excel_file)  