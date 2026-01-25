'''
Introduction:
    此程序用于AVL Excel文件的相关处理操作。
Usage:

Notes:

Revision log:
1.0.0 - 20260117: 初始版本，实现基本功能。
1.1.0 - 20260120: 添加compare_avl_sheets函数,用于对比AVL和AVL_Cmp两个sheet的数据
1.1.1 - 20260120: 修复compare_avl_sheets函数中AVL表E-T列跨分组匹配逻辑错误(顺序不同被标为红色的问题)
            正确的需求是只要AVL当前分组(如EF)在 AVL_Cmp 的所有分组集合中出现（无论顺序/位置），就应标为绿底，否则红底。
            同时测试代码输出文件名增加时间戳，避免覆盖。
1.2.0 - 20260121: 增加check_AVL_file()函数, 用于检查AVL Excel文件的有效性。
1.3.0 - 20260121: 增加get_SAP_Numbers_from_AVL_sheet()函数, 用于获取AVL表中的Part列表。
1.3.1 - 20260124: 因PyPika_CONNECT库更新,excel_mapping中索引需调整.
            




'''


# 版本号
# xx.yy.zz
# xx: 大版本，架构性变化
# yy: 功能性新增
# zz: Bug修复
__revision__ = '1.3.0'


import datetime
import os
import openpyxl
from openpyxl.styles import PatternFill
import copy

# Excel列与sql_result索引的对应关系
# sql result索引关系见third_party\PyPika_CONNECT\PyPika\PyPika_CONNECT.py中的FIELDS_SAPMaxDB， FIELDS_AccessDB
# 第一个值表示Excel的列，第二个值表示sql_result的索引，从0开始计数，None表示该列为空或不处理
excel_mapping = [
    ('A', None),        # 序号，特殊处理
    ('B', 2),           # SAP_Number
    ('C', 3),           # SAP_Description
    ('D', 34),          # detaildrawing 
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
    ('U', 35),          # AVL_Status
    ('V', 36),          # Editor
    ('W', 38),          # Technical Description, e.g. FCN
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

def copy_AVL_to_AVL_Cmp_In_UploadFile(db_search_result_file_path, upload_AVL_file_path):
    """
    复制db_search_result_file_path Excel 文件中的AVL工作表到upload_AVL_file_path的AVL_Cmp工作表。
    Args:
        db_search_result_file_path(str): 输入Excel文件路径
        upload_AVL_file_path (str): 输出Excel文件路径
    return:
        None: upload_AVL_file_path文件中新增AVL_Cmp工作表，内容与db_search_result_file_path文件中的AVL工作表相同
    """
    # 加载源文件和目标文件
    wb_source = openpyxl.load_workbook(db_search_result_file_path)
    wb_target = openpyxl.load_workbook(upload_AVL_file_path)

    # 获取AVL工作表
    ws_source_avl = wb_source["AVL"]

    # 在目标文件中创建AVL_Cmp工作表
    if "AVL_Cmp" in wb_target.sheetnames:
        ws_target_avl_cmp = wb_target["AVL_Cmp"]
        wb_target.remove(ws_target_avl_cmp)
    ws_target_avl_cmp = wb_target.create_sheet("AVL_Cmp")

    # 复制内容和样式
    for i, row in enumerate(ws_source_avl.iter_rows()):
        for j, cell in enumerate(row):
            new_cell = ws_target_avl_cmp.cell(row=i+1, column=j+1, value=cell.value)
            if cell.has_style:
                new_cell.font = copy.copy(cell.font)
                new_cell.border = copy.copy(cell.border)
                new_cell.fill = copy.copy(cell.fill)
                new_cell.number_format = cell.number_format
                new_cell.protection = copy.copy(cell.protection)
                new_cell.alignment = copy.copy(cell.alignment)

    # 复制合并单元格
    for merged_range in ws_source_avl.merged_cells.ranges:
        ws_target_avl_cmp.merge_cells(str(merged_range))

    # 保存目标文件
    wb_target.save(upload_AVL_file_path)    
    
def get_PCBA_Part_Numbers_from_BOM_Related_sheet(file_path):
    """
    获取BOM Related工作表中的PCBA Part Numbers列表。
    Args:
        file_path (str): 输入Excel文件路径
    return:
        PCBA_part_list (list): BOM Related工作表中的PCBA Part Numbers列表
    """
    PCBA_part_list = []
    try:
        wb = openpyxl.load_workbook(file_path)
        ws_bom = wb["BOM Related"]
        data_start_row = 3
        for row in range(data_start_row, ws_bom.max_row + 1):
            part_number = ws_bom[f"B{row}"].value
            if part_number is not None:
                PCBA_part_list.append(part_number)
        return PCBA_part_list
    except Exception as e:
        print(f"无法获取BOM Related工作表中的PCBA Part Numbers列表: {e}")
        return PCBA_part_list

def compare_avl_sheets(file_path, output_path):
    """
    比较AVL和AVL_Cmp两个sheet的数据，并根据比较结果进行标注。
    Args:
        param file_path (str): 输入Excel文件路径，包含AVL和AVL_Cmp两个sheet
        param output_path (str): 输出Excel文件路径，保存标注后的结果
    return:
        None: 结果保存至output_path文件中
    note:
        标注规则：
        1. 对于C/S/T列，直接比较两个sheet对应单元格的值：
           - 相同：绿色填充
           - 不同：红色填充
           - AVL有，AVL_Cmp无（无对应B列）：蓝色填充
        2. 对于E-T列，按分组（E&F, G&H, I&J, K&L, M&N, O&P, Q&R）进行跨行匹配比较：
           - 如果AVL的某组值在AVL_Cmp的任意组中存在且位置相同：绿色填充
           - 如果AVL的某组值在AVL_Cmp的任意组中存在但位置不同：红色填充
           - 如果AVL的某组值在AVL_Cmp的任意组中不存在：红色填充
        3. 对于AVL_Cmp表：
           - AVL_Cmp有，AVL无：整行橙色填充
           - AVL_Cmp有，AVL也有，但某组数据为新ordering code（即该组值在AVL中不存在）：该组单元格黄色填充
        函数由GPT-4.1协助编写和优化
        以下为prompt内容：
        ```我有一份Excel文件AVL_Cmp_Same_List_Example.xlsx
        里面有两张表AVL, AVL_Cmp
        他们的表格结构是一样的，第6列为表头
        从第7行开始就是数据
        两表B列的数据是一样的
        我需要使用python编写程序，对比两表每一行中C到T列的数据，
        以AVL表的B列数据为关联键，在AVL_Cmp表中检索到同一行数据
        1.若AVL_Cmp中有关联键对应的行
        对于C,S,T列，若两表同一单元格值相同，则在AVL表中设置为底色绿色，不同则为红色
        对于E至T列，相邻的两列为一组数据，如E、F列为一组数据，G、H列为一组数据，以此类推
        需要判断AVL表中每组数据的单元格值是否在AVL_Cmp表同一关联键行中存在(位置不是一一对应的，即AVL表EF列可能在GH列中对应，需要进行全检索)
        如果存在并相同，则在AVL表中标注为绿底；
        如果没有找到，则在AVL表中标注为红底;
        如果AVL_Cmp表中有，AVL表中没有，则在AVL_Cmp表中标注为黄底
        2.若AVL_Cmp中没有关联键对应的行
        则在AVL表中将该行标注为蓝底
        3.若AVL_Cmp中关联键在AVL表中没有数据
        则在AVL_Cmp表中将该行标注为橙底

        最后，填写图例信息
        在AVL表中，
        H1:绿底;内容：same
        I1:红底；内容：different
        J1:黄底;内容：AVL_Cmp New Ordering Code
        K1:蓝底;内容: Part to delete
        L1:橙底;内容: AVL_Cmp new part

        举例：
        AVL表中第7列数据为(Tab键隔开)
        1	2TFU901279U1001	SMD Cer.Cap 10uF 10% 16V 0805 X7R	""	Eyang	C0805X7R106K160NTH	FengHua	0805B106K160NT
        AVL_Cmp表中第7列数据为(Tab键隔开)
        1	2TFU901279U1001	SMD Cer.Cap 10uF 10% 16V 0805 X7R	""	FengHua	0805B106K160N3	Eyang	C0805X7R106K160NTH
    ```
    """
    # 定义填充颜色（RGB值）
    green_fill = PatternFill(start_color="00FF00", end_color="00FF00", fill_type="solid")  # 绿色：值相同/存在且匹配
    red_fill = PatternFill(start_color="FF0000", end_color="FF0000", fill_type="solid")    # 红色：值不同/存在但不匹配
    blue_fill = PatternFill(start_color="28A6EF", end_color="28A6EF", fill_type="solid")    # 浅蓝色：AVL有，AVL_Cmp无（无对应B列）
    yellow_fill = PatternFill(start_color="FFFF00", end_color="FFFF00", fill_type="solid")  # 黄色：AVL_Cmp有，AVL无
    orange_fill = PatternFill(start_color="FFA500", end_color="FFA500", fill_type="solid")  # 橙色：AVL_Cmp新part

    # 打开Excel文件28A6EF
    wb = openpyxl.load_workbook(file_path)
    # 获取两个工作表
    ws_avl = wb["AVL"]
    ws_cmp = wb["AVL_Cmp"]

    # 表头行是第6行，数据从第7行开始
    header_row = 6
    data_start_row = 7

    # ------------------- 步骤1：重新整理数据结构 -------------------
    # 存储AVL表数据：key=B列值, value={列名: 单元格值, "row_num": 行号}
    avl_data = {}
    # 存储AVL_Cmp表数据：
    # key=B列值, value={
    #   "row_num": 行号, 
    #   "col_data": {列名: 单元格值},
    #   "group_values": 所有分组值的集合（如{"E|F值", "G|H值"...}）
    # }
    cmp_data = {}
    # 所有Cmp表的分组值（全局）：用于跨分组匹配检查
    all_cmp_group_values = set()

    # E-T列分组定义（E&F, G&H...）
    groups = [("E", "F"), ("G", "H"), ("I", "J"), ("K", "L"), ("M", "N"), ("O", "P"), ("Q", "R")]

    # 遍历AVL表数据行
    for row in range(data_start_row, ws_avl.max_row + 1):
        b_value = ws_avl[f"B{row}"].value
        if b_value is None:
            continue
        row_data = {"row_num": row}
        # 读取C到T列的所有值
        for col in range(3, 21):  # C=3, T=20
            col_letter = openpyxl.utils.get_column_letter(col)
            row_data[col_letter] = ws_avl[f"{col_letter}{row}"].value
        avl_data[b_value] = row_data

    # 遍历AVL_Cmp表数据行（同时提取分组值）
    for row in range(data_start_row, ws_cmp.max_row + 1):
        b_value = ws_cmp[f"B{row}"].value
        if b_value is None:
            continue
        
        # 存储列数据
        col_data = {}
        for col in range(3, 21):
            col_letter = openpyxl.utils.get_column_letter(col)
            col_data[col_letter] = ws_cmp[f"{col_letter}{row}"].value
        
        # 提取当前行的所有分组值
        row_group_values = set()
        for col1, col2 in groups:
            val1 = col_data.get(col1, None)
            val2 = col_data.get(col2, None)
            group_val = f"{val1}|{val2}"
            row_group_values.add(group_val)
            all_cmp_group_values.add(group_val)  # 加入全局集合
        
        # 存入Cmp数据字典
        cmp_data[b_value] = {
            "row_num": row,
            "col_data": col_data,
            "group_values": row_group_values
        }

    # ------------------- 步骤2：修复AVL表标注逻辑 -------------------
    for b_key, avl_row in avl_data.items():
        avl_row_num = avl_row["row_num"]
        cmp_row = cmp_data.get(b_key, None)  # 获取对应Cmp行

        # ---------- 规则1：C/S/T列 直接对比值 ----------
        for col_letter in ["C", "S", "T"]:
            avl_cell_value = avl_row.get(col_letter, None)
            cmp_cell_value = cmp_row["col_data"].get(col_letter, None) if cmp_row else None

            avl_cell = ws_avl[f"{col_letter}{avl_row_num}"]
            if cmp_row is None:
                # 无对应B列 → 蓝色
                avl_cell.fill = blue_fill
            else:
                # 有对应B列，值相同绿，不同红
                avl_cell.fill = green_fill if (avl_cell_value == cmp_cell_value) else red_fill

        # ---------- 规则2：E-T列 跨分组匹配对比 ----------
        for col1, col2 in groups:
            # 获取AVL当前分组值
            avl_val1 = avl_row.get(col1, None)
            avl_val2 = avl_row.get(col2, None)
            avl_group_val = f"{avl_val1}|{avl_val2}"

            # 定位AVL单元格
            cell1 = ws_avl[f"{col1}{avl_row_num}"]
            cell2 = ws_avl[f"{col2}{avl_row_num}"]

            # 无对应B列 → 蓝色
            if cmp_row is None:
                cell1.fill = blue_fill
                cell2.fill = blue_fill
                continue

            # 只要分组值存在于Cmp当前行的所有分组中（无论顺序/位置），就绿底，否则红底
            if avl_group_val in cmp_row["group_values"]:
                cell1.fill = green_fill
                cell2.fill = green_fill
            else:
                cell1.fill = red_fill
                cell2.fill = red_fill

    # ------------------- 步骤3：修复AVL_Cmp表黄色/橙色标注逻辑 -------------------
    for b_key, cmp_row in cmp_data.items():
        cmp_row_num = cmp_row["row_num"]
        if b_key not in avl_data:
            # AVL_Cmp有，AVL无：整行橙色
            for col in range(3, 21):
                col_letter = openpyxl.utils.get_column_letter(col)
                ws_cmp[f"{col_letter}{cmp_row_num}"].fill = orange_fill
        else:
            # AVL_Cmp有，AVL也有，检查每组数据是否为新ordering code（AVL_Cmp有，AVL无）
            avl_row = avl_data[b_key]
            avl_groups = set()
            for col1, col2 in groups:
                avl_val1 = avl_row.get(col1, None)
                avl_val2 = avl_row.get(col2, None)
                avl_groups.add(f"{avl_val1}|{avl_val2}")
            for col1, col2 in groups:
                cmp_val1 = cmp_row["col_data"].get(col1, None)
                cmp_val2 = cmp_row["col_data"].get(col2, None)
                cmp_group_val = f"{cmp_val1}|{cmp_val2}"
                if cmp_group_val not in avl_groups:
                    ws_cmp[f"{col1}{cmp_row_num}"].fill = yellow_fill
                    ws_cmp[f"{col2}{cmp_row_num}"].fill = yellow_fill

    # ------------------- 步骤4：图例信息 -------------------
    ws_avl["H1"] = "same"
    ws_avl["H1"].fill = green_fill
    ws_avl["I1"] = "different"
    ws_avl["I1"].fill = red_fill
    ws_avl["J1"] = "AVL_Cmp New Ordering Code"
    ws_avl["J1"].fill = yellow_fill
    ws_avl["K1"] = "Part to delete"
    ws_avl["K1"].fill = blue_fill
    ws_avl["L1"] = "AVL_Cmp new part"
    ws_avl["L1"].fill = orange_fill

    # 保存结果
    wb.save(output_path)
    print(f"对比完成！结果已保存至: {output_path}")


AVL_MANUAL_REQUIRED_SHEETS = ["AVL", "AVL_Cmp"] # 手动整理的AVL对比所需工作表
AVL_AUTO_REQUIRED_SHEETS = ["AVL"] # 自动整理的AVL所需工作表
def check_AVL_file(file_path, required_sheets=["AVL", "AVL_Cmp"]):
    """
    检查AVL Excel文件的有效性。
    Args:
        param file_path (str): 输入Excel文件路径
    return:
        bool: 文件有效返回True，否则返回False
    """
    try:
        wb = openpyxl.load_workbook(file_path)
        # 检查是否包含所需的工作表
        for sheet in required_sheets:
            if sheet not in wb.sheetnames:
                print(f"缺少必要的工作表: {sheet}")
                return False
        return True
    except Exception as e:
        print(f"无法打开文件或文件格式错误: {e}")
        return False

def get_SAP_Numbers_from_AVL_sheet(file_path):
    """
    获取AVL表中的Part列表。
    Args:
        param file_path (str): 输入Excel文件路径
    return:
        list: AVL表中的Part列表
    """
    part_list = []
    try:
        wb = openpyxl.load_workbook(file_path)
        ws_avl = wb["AVL"]
        data_start_row = 7
        for row in range(data_start_row, ws_avl.max_row + 1):
            b_value = ws_avl[f"B{row}"].value
            if b_value is not None:
                part_list.append(b_value)
        return part_list
    except Exception as e:
        print(f"无法获取AVL表中的Part列表: {e}")
        return part_list

# 主程序执行
if __name__ == "__main__":
    # 获取当前脚本所在目录
    script_dir = os.path.dirname(os.path.abspath(__file__))
    
    # 基于脚本目录构建文件路径
    input_file = os.path.join(script_dir, "AVL_Cmp_Same_List_Example.xlsx")
    # 输出文件名要增加日期时间戳，避免覆盖
    timestamp = datetime.datetime.now().strftime("%Y%m%d_%H%M%S")
    output_file = os.path.join(script_dir, f"AVL_Cmp_Same_List_Example_compared_{timestamp}.xlsx")  
    
    # compare_avl_sheets(input_file, output_file)

    # debug: 获取BOM Related工作表中的PCBA Part Numbers列表
    PCBA_Part_List = get_PCBA_Part_Numbers_from_BOM_Related_sheet(input_file)
    print("BOM Related工作表中的PCBA Part Numbers列表:")
    print(PCBA_Part_List)

