import pypyodbc
import pandas as pd
from typing import Dict, List
import warnings
warnings.filterwarnings('ignore')  # 忽略无关警告

# 你提供的原始字段/表列表（无需修改）
FIELDS_SAPMaxDB: List[str] = [
    "PartNumber", "value_1", "SAP_Number", "SAP_Description", "status", "parttype",
    "manufact_1", "manufact_partnum_1", "datasheet_1",
    "manufact_2", "manufact_partnum_2", "datasheet_2",
    "manufact_3", "manufact_partnum_3", "datasheet_3",
    "manufact_4", "manufact_partnum_4", "datasheet_4",
    "manufact_5", "manufact_partnum_5", "datasheet_5",
    "manufact_6", "manufact_partnum_6", "datasheet_6",
    "manufact_7", "manufact_partnum_7", "datasheet_7",
    "scm_symbol", "pcb_footprint", "alt_symbols",
    "mounttechn", "ad_symbol", "ad_footprint", "ad_alt_footprint", "detaildrawing",
    "Status", "Editor", "US_technology", "TechDescription"
]

FIELDS_AccessDB: List[str] = [
    "PartNumber", "value", "SAP_Number", "SAP_Description", "status", "parttype",
    "[manufact 1]", "[manufact partnum 1]", "[datasheet 1]",
    "[manufact 2]", "[manufact partnum 2]", "[datasheet 2]",
    "[manufact 3]", "[manufact partnum 3]", "[datasheet 3]",
    "[manufact 4]", "[manufact partnum 4]", "[datasheet 4]",
    "[manufact 5]", "[manufact partnum 5]", "[datasheet 5]",
    "[manufact 6]", "[manufact partnum 6]", "[datasheet 6]",
    "[manufact 7]", "[manufact partnum 7]", "[datasheet 7]",
    "scm_symbol", "pcb_footprint", "alt_symbols",
    "mounttechn", "ad_symbol", "ad_footprint", "ad_alt_footprint", "detaildrawing",
    "STATUS", "EDITOR", "US_TECHNOLOGY", "TECHDESCRIPTION"
]

TABLES_SAPMaxDB: List[str] = [
    "CAPACITORS", 
    # "CONNECTORS", "CONVERTERS", "DIODES", "ICS_ANALOG",
    # "ICS_DIGITAL", "MAGNETICS", "MECHPARTS", "MEMORY", "MISCPARTS",
    # "OPTO", "OP_AMPS", "OSCILLATORS", "REGULATORS", "RELAYS",
    # "RESISTORS", "SENSORS", "SWITCHES", "TRANSFORMERS", "TRANSISTORS",
    # "VARISTORS"
]

TABLES_AccessDB: List[str] = [
    "[01-Capacitors]", 
    # "[02-Resistors]", "[03-Varistors]", "[04-Transistors]", "[05-Diodes]",
    # "[06-ICs_digital]", "[07-Memory]", "[08-ICs_analog]", "[09-Regulators]", "[10-Converters]",
    # "[11-OP_Amps]", "[12-Magnetics]", "[13-Transformers]", "[14-Opto]", "[15-Oscillators]",
    # "[16-Connectors]", "[17-Relays]", "[18-Sensors]", "[19-Switches]", "[20-MechParts]",
    # "[21-MiscParts]"
]

# ===================== 【核心配置 - 需根据实际环境修改】=====================
# 1. 表名一一映射：SAPMaxDB表名 → AccessDB表名（已按你提供的列表对应，无需修改）
TABLE_MAPPING: Dict[str, str] = dict(zip(
    TABLES_SAPMaxDB,
    TABLES_AccessDB
))

# 2. 字段一一映射：SAPMaxDB字段 → AccessDB字段（核心：解决字段名不一致问题）
FIELD_MAPPING: Dict[str, str] = dict(zip(
    FIELDS_SAPMaxDB,
    FIELDS_AccessDB
))

# 3. ODBC数据库连接配置（关键：替换为你的实际DSN/用户名/密码）
DB_CONFIG = {
    'sap': {
        'dsn': 'CIS_Local',  # 替换为SAPMaxDB的ODBC数据源名称
        'user': 'LIMBAS2USER',  # 无则留空
        'password': 'LIMBASREAD',  # 无则留空
    },
    'access': {
        'dsn': 'CIS_PartLib_P_64',  # 替换为AccessDB的ODBC数据源名称
        'user': 'cadence_port',
        'password': 'Cadence_CIS.3',
    }
}

# 4. 对比核心规则
UNIQUE_KEY = 'partnumber'  # 唯一标识字段（与配置中字段名保持一致）, 输出时全部为小写
EXCEL_OUTPUT_PATH = 'SAP_Access数据对比结果.xlsx'  # Excel输出路径
# ==========================================================================


# ===================== 数据库操作函数（pypyodbc专属）=====================
def get_db_connection(db_type: str) -> pypyodbc.Connection:
    """
    获取pypyodbc数据库连接（sap/access）
    :param db_type: 数据库类型，可选'sap'/'access'
    :return: pypyodbc连接对象
    """
    if db_type not in DB_CONFIG:
        raise ValueError(f"数据库类型仅支持'sap'和'access'，当前传入：{db_type}")
    
    config = DB_CONFIG[db_type]
    # 构造pypyodbc ODBC连接字符串（极简版，适配绝大多数ODBC配置）
    conn_str = f"DSN={config['dsn']};"
    if config['user']:
        conn_str += f"UID={config['user']};"
    if config['password']:
        conn_str += f"PWD={config['password']};"
    
    try:
        # pypyodbc连接：autocommit=True避免事务锁定
        conn = pypyodbc.connect(conn_str, autocommit=True)
        print(f"✅ {db_type.upper()}数据库连接成功（pypyodbc）")
        return conn
    except pypyodbc.Error as e:
        raise ConnectionError(f"❌ {db_type.upper()}数据库连接失败：{str(e)}")

def read_specified_fields(conn: pypyodbc.Connection, table_name: str, fields: List[str]) -> pd.DataFrame:
    """
    读取指定表的**指定字段**数据（适配特殊字段名/表名）
    :param conn: pypyodbc连接对象
    :param table_name: 表名（支持[]包裹的特殊表名）
    :param fields: 需查询的字段列表（支持[]包裹的特殊字段名）
    :return: 包含指定字段的DataFrame，空表返回空DataFrame
    """
    try:
        # 拼接字段字符串：字段间用,分隔
        fields_str = ", ".join(fields)
        # 构造查询SQL：仅查询指定字段，提升效率
        sql = f"SELECT {fields_str} FROM {table_name}"
        # pandas读取pypyodbc数据，自动适配字段名
        df = pd.read_sql(sql, conn)
        
        # 关键处理：唯一标识字段转为字符串并去空格，避免数字/字符串对比错误
        if UNIQUE_KEY in df.columns:
            df[UNIQUE_KEY] = df[UNIQUE_KEY].astype(str).str.strip()
            # 去重：保留唯一的PartNumber（避免重复数据干扰对比）
            df = df.drop_duplicates(subset=[UNIQUE_KEY], keep='first')
        
        print(f"✅ 读取{table_name}成功，字段数：{len(fields)}，数据量：{len(df)}条")
        return df
    except pypyodbc.Error as e:
        raise Exception(f"❌ 读取表{table_name}失败：{str(e)}")


# ===================== 数据对比核心函数（按字段映射）=====================
def compare_tables_by_mapping(sap_df: pd.DataFrame, access_df: pd.DataFrame, sap_table: str) -> pd.DataFrame:
    """
    按字段映射关系，对比SAP和Access表数据（以PartNumber为唯一键）
    :param sap_df: SAPMaxDB表的DataFrame（含FIELDS_SAPMaxDB字段）
    :param access_df: AccessDB表的DataFrame（含FIELDS_AccessDB字段）
    :param sap_table: SAP表名（用于日志）
    :return: 带差异标记的对比结果DataFrame
    """
    # 校验唯一标识字段是否存在
    for df, db_name in [(sap_df, 'SAPMaxDB'), (access_df, 'AccessDB')]:
        if UNIQUE_KEY not in df.columns:
            raise ValueError(f"❌ {db_name}表{sap_table}缺少唯一标识字段{UNIQUE_KEY}")
        if len(df) == 0:
            raise Exception(f"❌ {db_name}表{sap_table}无数据，无法对比")
    
    # 对AccessDF重命名：按字段映射将Access字段名改为SAP字段名，实现字段对齐
    # 反向映射：Access原字段 → SAP标准字段
    access_field_rename = {v: k for k, v in FIELD_MAPPING.items()}
    access_df_renamed = access_df.rename(columns=access_field_rename)
    
    # 合并两个表：外连接（保留双方所有PartNumber，无匹配则显示NaN）
    # 仅保留映射后的共同字段（即SAP的标准字段）
    merge_df = pd.merge(
        sap_df,
        access_df_renamed,
        on=UNIQUE_KEY,
        how='outer',
        suffixes=('_SAP', '_Access')  # 同名字段添加库标识后缀
    )
    
    # 逐字段对比：生成字段级差异标记
    compare_flags = []
    for sap_field in FIELDS_SAPMaxDB:
        if sap_field == UNIQUE_KEY:
            continue  # 唯一键无需对比
        sap_col = f"{sap_field}_SAP"
        access_col = f"{sap_field}_Access"
        # 对比规则：处理NaN/空值，统一转为字符串后对比，去除首尾空格
        sap_vals = merge_df[sap_col].fillna('').astype(str).str.strip()
        access_vals = merge_df[access_col].fillna('').astype(str).str.strip()
        # 生成差异标记：True=一致，False=不一致
        merge_df[f"差异_{sap_field}"] = (sap_vals == access_vals)
        compare_flags.append(f"差异_{sap_field}")
    
    # 生成整体差异标记：只要有一个字段不一致，即为差异记录
    merge_df['整体差异标记'] = ~merge_df[compare_flags].all(axis=1)
    merge_df['整体差异标记'] = merge_df['整体差异标记'].map({
        True: '❌ 存在差异',
        False: '✅ 完全一致'
    })
    
    # 处理无匹配的PartNumber：标记来源
    merge_df[UNIQUE_KEY] = merge_df[UNIQUE_KEY].fillna('【无匹配PartNumber】')
    # SAP无此记录
    merge_df.loc[merge_df[f"{UNIQUE_KEY}_SAP"].isna(), UNIQUE_KEY] = merge_df[f"{UNIQUE_KEY}_Access"] + "【Access独有】"
    # Access无此记录
    merge_df.loc[merge_df[f"{UNIQUE_KEY}_Access"].isna(), UNIQUE_KEY] = merge_df[f"{UNIQUE_KEY}_SAP"] + "【SAP独有】"
    
    # 调整列顺序：先PartNumber → 整体差异标记 → SAP字段 → Access字段 → 字段级差异标记
    col_order = [
        UNIQUE_KEY,
        '整体差异标记'
    ] + [f"{f}_SAP" for f in FIELDS_SAPMaxDB if f != UNIQUE_KEY] + \
      [f"{f}_Access" for f in FIELDS_SAPMaxDB if f != UNIQUE_KEY] + \
      compare_flags
    # 过滤有效列（避免因空表导致的列缺失）
    col_order = [col for col in col_order if col in merge_df.columns]
    final_df = merge_df[col_order]
    
    # 统计差异数
    diff_count = (final_df['整体差异标记'] == '❌ 存在差异').sum()
    print(f"✅ 表{sap_table}对比完成，总记录数：{len(final_df)}，差异记录数：{diff_count}")
    return final_df

# ===================== 主执行程序（批量对比+多sheet导出）=====================
def main():
    # 1. 建立数据库连接（pypyodbc）
    try:
        sap_conn = get_db_connection('sap')
        access_conn = get_db_connection('access')
    except Exception as e:
        print(f"程序终止：{str(e)}")
        return
    
    # 2. 创建Excel写入器（多sheet支持，engine=openpyxl）
    try:
        with pd.ExcelWriter(EXCEL_OUTPUT_PATH, engine='openpyxl') as writer:
            # 3. 按表名映射批量执行对比
            for sap_table, access_table in TABLE_MAPPING.items():
                print(f"\n========== 开始对比：SAP[{sap_table}] <-> Access[{access_table}] ==========")
                try:
                    # 读取指定表的指定字段数据
                    access_df = read_specified_fields(access_conn, access_table, FIELDS_AccessDB)
                    sap_df = read_specified_fields(sap_conn, sap_table, FIELDS_SAPMaxDB)
                    
                    # 空表跳过
                    if len(sap_df) == 0 and len(access_df) == 0:
                        print(f"⚠️  两个表均为空，跳过对比")
                        continue
                    
                    # 执行精准对比
                    compare_result_df = compare_tables_by_mapping(sap_df, access_df, sap_table)
                    
                    # 写入Excel：sheet名取SAP表名，避免特殊字符
                    compare_result_df.to_excel(writer, sheet_name=sap_table, index=False)
                    print(f"✅ 表{sap_table}对比结果已写入Excel")
                
                except Exception as e:
                    print(f"⚠️  表{sap_table}对比失败，跳过：{str(e)}")
                    continue
        
        print(f"\n========== 所有表对比完成 ==========")
        print(f"📊 最终对比结果已导出至：{EXCEL_OUTPUT_PATH}")
    
    except Exception as e:
        print(f"❌ Excel导出失败：{str(e)}")
    finally:
        # 4. 强制关闭数据库连接，释放资源
        sap_conn.close()
        access_conn.close()
        print(f"✅ 数据库连接已全部关闭")

# 程序入口
if __name__ == '__main__':
    main()