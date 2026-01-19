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


import os
import openpyxl
from openpyxl.styles import PatternFill

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

def first_write_AVL_to_excel(template_file, sql_result, Multi_PCBA_Part_info_list, output_excel_file):
    """
    首次将AVL数据写入Excel文件。
    Args:
        param template_file (str): 模板文件路径
        param sql_result (list): 从数据库查询得到的AVL数据列表
        param Multi_PCBA_Part_info_list (list): 从PLM获取的PCBA物料信息列表
        param output_excel_file (str): 输出Excel文件路径
    return: 
        None
    """

    # 加载模板文件
    wb = openpyxl.load_workbook(template_file)
    # AVL sheet保存数据
    sheet_avl = wb['AVL']
    # AVL sheet保存数据
    for row_idx, row_data in enumerate(sql_result, start=7):  # Excel第7行开始
        if len(row_data) > 4: # 数据有效， 表示查询到了数据
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
        else: # 找不到数据，只填写序号、SAP_Number、SAP_Description
            # A 列为序号
            sheet_avl[f"A{row_idx}"] = row_idx - 6
            # B 列为SAP_Number
            sheet_avl[f"B{row_idx}"] = row_data[2]
            # C 列为SAP_Description
            sheet_avl[f"C{row_idx}"] = row_data[3]

    # BOM Related sheet保存数据
    sheet_bom = wb['BOM Related']
    for row_idx, row_data in enumerate(Multi_PCBA_Part_info_list, start=3):  # Excel
        if row_data != []:
            print(row_idx, row_data)
            sheet_bom.cell(row=row_idx, column=2).value = row_data[0]  # PCBA Part Number
            sheet_bom.cell(row=row_idx, column=3).value = row_data[3]  # PCBA Part Description
    # 保存输出文件
    wb.save(output_excel_file)  

    


def compare_avl_sheets(file_path, output_path):
    # 定义填充颜色（RGB值）
    green_fill = PatternFill(start_color="00FF00", end_color="00FF00", fill_type="solid")  # 绿色：值相同
    red_fill = PatternFill(start_color="FF0000", end_color="FF0000", fill_type="solid")    # 红色：值不同/有差异
    blue_fill = PatternFill(start_color="0000FF", end_color="0000FF", fill_type="solid")    # 蓝色：AVL有，AVL_Cmp无
    yellow_fill = PatternFill(start_color="FFFF00", end_color="FFFF00", fill_type="solid")  # 黄色：AVL_Cmp有，AVL无

    # 打开Excel文件
    wb = openpyxl.load_workbook(file_path)
    # 获取两个工作表
    ws_avl = wb["AVL"]
    ws_cmp = wb["AVL_Cmp"]

    # 步骤1：整理数据 - 以B列为键，存储每行数据（行号: 列值字典）
    # 表头行是第6行，数据从第7行开始
    header_row = 6
    data_start_row = 7

    # 存储AVL表数据：key=B列值, value={列名: 单元格值, "row_num": 行号}
    avl_data = {}
    # 存储AVL_Cmp表数据：key=B列值, value={列名: 单元格值, "row_num": 行号}
    cmp_data = {}

    # 遍历AVL表数据行
    for row in range(data_start_row, ws_avl.max_row + 1):
        b_value = ws_avl[f"B{row}"].value
        if b_value is None:
            continue  # 跳过B列为空的行
        row_data = {"row_num": row}
        # 读取C到T列的所有值
        for col in range(3, 21):  # C=3, T=20
            col_letter = openpyxl.utils.get_column_letter(col)
            row_data[col_letter] = ws_avl[f"{col_letter}{row}"].value
        avl_data[b_value] = row_data

    # 遍历AVL_Cmp表数据行
    for row in range(data_start_row, ws_cmp.max_row + 1):
        b_value = ws_cmp[f"B{row}"].value
        if b_value is None:
            continue  # 跳过B列为空的行
        row_data = {"row_num": row}
        # 读取C到T列的所有值
        for col in range(3, 21):  # C=3, T=20
            col_letter = openpyxl.utils.get_column_letter(col)
            row_data[col_letter] = ws_cmp[f"{col_letter}{row}"].value
        cmp_data[b_value] = row_data

    # 步骤2：逐行对比并设置底色
    # 先处理AVL表的单元格标色（核心对比逻辑）
    for b_key, avl_row in avl_data.items():
        avl_row_num = avl_row["row_num"]
        # 获取对应的Cmp行数据（B列值匹配）
        cmp_row = cmp_data.get(b_key, None)

        # ---------- 规则1：C/S/T列 直接对比值 ----------
        for col_letter in ["C", "S", "T"]:
            avl_cell_value = avl_row.get(col_letter, None)
            cmp_cell_value = cmp_row.get(col_letter, None) if cmp_row else None

            # 定位AVL表的单元格
            avl_cell = ws_avl[f"{col_letter}{avl_row_num}"]
            # 对比值：相同标绿，不同标红
            if avl_cell_value == cmp_cell_value:
                avl_cell.fill = green_fill
            else:
                avl_cell.fill = red_fill

        # ---------- 规则2：E-T列 按相邻两列分组对比 ----------
        # 分组：E&F, G&H, I&J, K&L, M&N, O&P, Q&R (覆盖E-T列)
        groups = [("E", "F"), ("G", "H"), ("I", "J"), ("K", "L"), ("M", "N"), ("O", "P"), ("Q", "R")]
        for col1, col2 in groups:
            # 获取AVL的分组值（拼接为字符串，方便对比）
            avl_val1 = avl_row.get(col1, None)
            avl_val2 = avl_row.get(col2, None)
            avl_group_val = f"{avl_val1}|{avl_val2}"

            # 获取Cmp的分组值（如果有匹配的B列）
            if cmp_row:
                cmp_val1 = cmp_row.get(col1, None)
                cmp_val2 = cmp_row.get(col2, None)
                cmp_group_val = f"{cmp_val1}|{cmp_val2}"
            else:
                cmp_group_val = None

            # 定位AVL表的两个单元格
            cell1 = ws_avl[f"{col1}{avl_row_num}"]
            cell2 = ws_avl[f"{col2}{avl_row_num}"]

            # 判断并标色
            if cmp_row is None:
                # AVL有，AVL_Cmp无 → 蓝色
                cell1.fill = blue_fill
                cell2.fill = blue_fill
            else:
                if avl_group_val == cmp_group_val:
                    # 分组值完全相同 → 绿色
                    cell1.fill = green_fill
                    cell2.fill = green_fill
                elif avl_val1 is not None or avl_val2 is not None:
                    # 分组值有差异 → 红色
                    cell1.fill = red_fill
                    cell2.fill = red_fill

    # 步骤3：处理AVL_Cmp有但AVL无的情况（标黄色）
    for b_key, cmp_row in cmp_data.items():
        cmp_row_num = cmp_row["row_num"]
        # 如果AVL中没有该B列值，说明AVL_Cmp有但AVL无 → 标黄色
        if b_key not in avl_data:
            # E-T列分组标黄
            groups = [("E", "F"), ("G", "H"), ("I", "J"), ("K", "L"), ("M", "N"), ("O", "P"), ("Q", "R")]
            for col1, col2 in groups:
                cell1 = ws_cmp[f"{col1}{cmp_row_num}"]
                cell2 = ws_cmp[f"{col2}{cmp_row_num}"]
                cell1.fill = yellow_fill
                cell2.fill = yellow_fill
            # C/S/T列也标黄（因为AVL无对应值）
            for col_letter in ["C", "S", "T"]:
                cell = ws_cmp[f"{col_letter}{cmp_row_num}"]
                cell.fill = yellow_fill

    # 保存结果到新文件
    wb.save(output_path)
    print(f"对比完成！结果已保存至: {output_path}")

# 主程序执行
if __name__ == "__main__":
    # 获取当前脚本所在目录
    script_dir = os.path.dirname(os.path.abspath(__file__))
    
    # 基于脚本目录构建文件路径
    input_file = os.path.join(script_dir, "AVL_Cmp_Same_List_Example.xlsx")
    output_file = os.path.join(script_dir, "AVL_Cmp_Same_List_Example_compared.xlsx")
    
    compare_avl_sheets(input_file, output_file)
