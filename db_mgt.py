'''
Introduction:
This module provides a set of classes and functions to manage database connections and queries for different database systems. 
It supports multiple database types, including ODBC connections for local and online databases.

Revision History:
1.0.0 - 标注版本号的第一个版本, 在进行代码重构之前的版本.
2.0.0 - 20260108:使用PyPika子模块进行代码重构，优化数据库连接管理。
        相关文件: db_mgt.py, third_party/PyPika_CONNECT/PyPika/PyPika_CONNECT.py
        更新函数: fetch()
        功能影响: 
        1.单PartType搜索输出列表跟All PartType时一样
        2.All PartType搜索各个条件可以同时使用为AND关系
2.1.0 - 20260108: 新增过滤条件“Description”"techdescription" "editor"到fetch函数中。
        注意:
            SAPMaxDB中Editor字段基本都为空值,检查同一物料的AccessDB却是有值,如CAP_1630物料,导致搜索结果不一致。
            这个问题应该是数据库后台问题,需要反馈。
2.2.0 - 20260124: 更改DBList中四个数据库的描述名称, 使其更清晰易懂。
        修改后：
            第一个为直连DESTO Cancdence CIS数据库的ODBC连接
            第二个为直连CNILG服务器上Access数据库的ODBC连接(AD共用)
            第三个为CNILG服务器上Access数据库的文件连接
            第四个为CNILX服务器上Access数据库的文件连接
        所以第一个使用SAPMaxDB连接,后三个使用AccessDB连接。

'''


# 版本号
# xx.yy.zz
# xx: 大版本，架构性变化
# yy: 功能性新增
# zz: Bug修复
__version__ = "2.2.0"


# 导入子模块
# 直接按目录结构导入，无需sys.path
# 方法一:
# from third_party.PyPika_CONNECT.PyPika.PyPika_CONNECT import *
# print("PyPika_CONNECT Version:", __version__)
# 方法二:
import third_party.PyPika_CONNECT.PyPika.PyPika_CONNECT as PyPika_CONNECT
# print("Imported PyPika_CONNECT Version:", PyPika_CONNECT.__version__)

import pypyodbc



# windows parameter
# Database List can be used.
DBList = ['01-Cadence CIS DB(ODBC, DESTO)', 
          '02-Altium Access DB(ODBC, CNILG)', 
          '03-Access DB(File in CNILG)',
          '04-Access DB(File in CNILX)']


# Part Type List for DB: '01-Cadence CIS DB(ODBC, DESTO)'
PartTypeList_CONNECT = [
('---All----'),
('CAPACITORS'), 
('CONNECTORS'), 
('CONVERTERS'), 
('DIODES'), 
('ICS_ANALOG'), 
('ICS_DIGITAL'), 
('MAGNETICS'), 
('MECHPARTS'), 
('MEMORY'), 
('MISCPARTS'), 
('OPTO'), 
('OP_AMPS'), 
('OSCILLATORS'), 
('PCB'), 
('REGULATORS'), 
('RELAYS'), 
('RESISTORS'), 
('SENSORS'), 
('SHAPES'), 
('SOFTWARE'), 
('SWITCHES'), 
('TITLEBLOCK'), 
('TRANSFORMERS'), 
('TRANSISTORS'), 
('VARISTORS')]

PartTypeList_CONNECT_4All_Search = [
('CAPACITORS'), 
('CONNECTORS'), 
('CONVERTERS'), 
('DIODES'), 
('ICS_ANALOG'), 
('ICS_DIGITAL'), 
('MAGNETICS'), 
('MECHPARTS'), 
('MEMORY'), 
('MISCPARTS'), 
('OPTO'), 
('OP_AMPS'), 
('OSCILLATORS'), 
('REGULATORS'), 
('RELAYS'), 
('RESISTORS'), 
('SENSORS'), 
('SWITCHES'), 
('TRANSFORMERS'), 
('TRANSISTORS'), 
('VARISTORS')]

# Part Type list for DB: '02-Altium Access DB(ODBC, CNILG)'
PartTypeList_Access = [
 ('---All----'),
('01-Capacitors'),
 ('02-Resistors'),
 ('03-Varistors'),
 ('04-Transistors'),
 ('05-Diodes'),
 ('06-ICs_digital'),
 ('07-Memory'),
 ('08-ICs_analog'),
 ('09-Regulators'),
 ('10-Converters'),
 ('11-OP_Amps'),
 ('12-Magnetics'),
 ('13-Transformers'),
 ('14-Opto'),
 ('15-Oscillators'),
 ('16-Connectors'),
 ('17-Relays'),
 ('18-Sensors'),
 ('19-Switches'),
 ('20-MechParts'),
 ('21-MiscParts'),
 ('98-Shapes')
]
PartTypeList_Access_4All_Search = [
('01-Capacitors'),
 ('02-Resistors'),
 ('03-Varistors'),
 ('04-Transistors'),
 ('05-Diodes'),
 ('06-ICs_digital'),
 ('07-Memory'),
 ('08-ICs_analog'),
 ('09-Regulators'),
 ('10-Converters'),
 ('11-OP_Amps'),
 ('12-Magnetics'),
 ('13-Transformers'),
 ('14-Opto'),
 ('15-Oscillators'),
 ('16-Connectors'),
 ('17-Relays'),
 ('18-Sensors'),
 ('19-Switches'),
 ('20-MechParts'),
 ('21-MiscParts')
]

# Manufacture Part Number List for the sql
MftPartNumList_Access = ['[manufact partnum 1]','[manufact partnum 2]','[manufact partnum 3]','[manufact partnum 4]','[manufact partnum 5]','[manufact partnum 6]','[manufact partnum 7]']
MftPartNumList_SAPMax = ['manufact_partnum_1','manufact_partnum_2','manufact_partnum_3','manufact_partnum_4','manufact_partnum_5','manufact_partnum_6','manufact_partnum_7']

# Database control class
class Database:
    def __ini__(self):
        # 初始化不需要创建任务东西
        pass
    
    def defaul(self,dbindex):
        # template
        # 01-CONNECT Local(ODBC)
        if dbindex == 0:
            pass
        # 02-Access Online(ODBC)
        elif dbindex == 1:
            pass
        # 03-P Disk Access
        elif dbindex == 2:
            pass
        # 04-CONNECT DESTO(ODBC)
        elif dbindex == 3:
            pass
   
    def openDB(self, dbindex, dblist, app):
        """Open a database connection.

        Args:
            dbindex (int): The index of the database to connect to.
            dblist (list): The list of database names.  
            app (_type_): Flask app for logging.

        Returns:
            _type_: True if the connection is successful, False otherwise.
        """
        # 01-Cadence CIS DB(ODBC, DESTO)
        if dbindex == 0:
            connStr = "DSN=CIS_Local;Uid=LIMBAS2USER;Pwd=LIMBASREAD;"
            # connStr = "DSN=CONNECT Partslib V2;Uid=LIMBAS2USER;Pwd=LIMBASREAD;"
            print(dblist[dbindex])
        # 02-Altium Access DB(ODBC, CNILG)
        elif dbindex == 1:
            connStr = "DSN=CIS_PartLib_P_64;Uid=cadence_port;Pwd=Cadence_CIS.3;"
            print(dblist[dbindex])
        # 03-Access DB(File in CNILG)
        elif dbindex == 2:
            connStr = r"Driver={Microsoft Access Driver (*.mdb, *.accdb)};Dbq=\\cn-s-lns050b.cn.abb.com\orcad$\DESTODATABASE\Cadence\CIS_DB\CIS_PartLib.mdb;SystemDB=\\cn-s-lns050b.cn.abb.com\orcad$\DESTODATABASE\Cadence\CIS_DB\CIS_PartLib.mdw;Uid=cadence_port;Pwd=Cadence_CIS.3;"
            print(dblist[dbindex])
        # 04-Access DB(File in CNILX)
        elif dbindex == 3:
            connStr = r"Driver={Microsoft Access Driver (*.mdb, *.accdb)};Dbq=\\CN-S-APPC007P\01_EleTeam\Cadence\CIS_DB\CIS_PartLib.mdb;SystemDB=\\CN-S-APPC007P\01_EleTeam\Cadence\CIS_DB\CIS_PartLib.mdw;Uid=cadence_port;Pwd=Cadence_CIS.3;"
            print(dblist[dbindex])

        # 连接数据库
        try:
            self.conn = pypyodbc.connect(connStr, timeout=20, readonly=True)
            self.cursor = self.conn.cursor()
            print("Connect DB success!")
            return True
        except Exception as e:
            print("Cannot Connect to DB!!\n")
            app.logger.error('Cannot Connect to DB!!\n Error:%s',e)
            # print(e)
            return False
    
    def listTable(self, dbindex):    
        """List all tables in the database.

        Args:
            dbindex (int): The index of the database to list tables from.

        Returns:
            list: A list of table names.
        """
        # get the table list
        # 01-CONNECT Local(ODBC)
        if dbindex == 0 :
            # SAPMaxDB数据库获取表名的SQL语句
            sql_listTable = "select table_name from all_tables"
        # 02-Access Online(ODBC) and 03-P Disk Access
        elif dbindex == 1 or dbindex == 2 or dbindex == 3:
            # Access数据库获取表名的SQL语句
            sql_listTable = "SELECT NAME FROM MSYSOBJECTS WHERE TYPE=1 AND FLAGS=0;"
        self.cursor.execute(sql_listTable)
        table_list = self.cursor.fetchall()
        print(table_list)
        return table_list

    def fetch(self, tableName, dbindex, PartNo_Searchby, SAPNo_Searchby, PartValue_Searchby, MfcPartNum_Searchby, Description_Searchby, TechDescription_Searchby, Editor_Searchby):
        final_sql = ''
        # Determine DB_Type
        # 01-CONNECT Local(ODBC)
        # SAPMaxDB数据库获取表名的SQL语句
        if dbindex == 0 :
            DB_Type = "SAPMaxDB"
        # 02-Altium Access DB(ODBC, CNILG) and 03-Access DB(File in CNILG) and 04-Access DB(File in CNILX)
        # AccessDB数据库获取表名的SQL语句
        elif dbindex == 1 or dbindex == 2 or dbindex == 3: 
            DB_Type = "AccessDB"
        # print("Database Type:", DB_Type)

        # 仅需要单独判断是搜索某个表还是所有表
        # if not search all, search specified table
        if tableName != '---All----':
            TABLES = []
            # Determine TABLES and FIELDS based on DB_Type
            if DB_Type == "AccessDB":
                # AccessDB
                TABLES.append(f"[{tableName}]")
                FIELDS = PyPika_CONNECT.FIELDS_AccessDB
            else:
                # SAPMaxDB
                TABLES.append(f"{tableName}")
                FIELDS = PyPika_CONNECT.FIELDS_SAPMaxDB
        # serach all table
        else:          
            # Determine TABLES and FIELDS based on DB_Type
            if DB_Type == "AccessDB":
                # AccessDB
                TABLES = PyPika_CONNECT.TABLES_AccessDB
                FIELDS = PyPika_CONNECT.FIELDS_AccessDB
            else:
                # SAPMaxDB
                TABLES = PyPika_CONNECT.TABLES_SAPMaxDB
                FIELDS = PyPika_CONNECT.FIELDS_SAPMaxDB
            
        # 生成过滤条件
        FILTER_CONDITIONS = PyPika_CONNECT.generate_filter_conditions(
            DB_Type=DB_Type,
            PartNo_Searchby=PartNo_Searchby,
            SAPNo_Searchby=SAPNo_Searchby,
            PartValue_Searchby=PartValue_Searchby,
            MfcPartNum_Searchby=MfcPartNum_Searchby,
            Description_Searchby=Description_Searchby,
            TechDescription_Searchby=TechDescription_Searchby,
            Editor_Searchby=Editor_Searchby
        )

        # 生成最终SQL
        final_sql = PyPika_CONNECT.build_final_sql(
            tables=TABLES,
            fields=FIELDS,
            filter_conditions=FILTER_CONDITIONS,
            order_by_field="PartNumber",
            order="ASC"
        )
        # print("Generated SQL:\n", final_sql)
        self.cursor.execute(final_sql)
        columnNameList = [column[0] for column in self.cursor.description]
        sql_result = self.cursor.fetchall()
        # print(sql_result)
        return sql_result, columnNameList
        
    def openAcc(self):
        MDB=r"C:\inetpub\wwwroot\db\#PrJRcd.mdb"    # 此目录需要特殊权限才能访问，调试或运行时请注意，权限不足会报错，如pypyodbc.ProgrammingError: ('42000', '[42000] [Microsoft][ODBC Microsoft Access Driver] Not a valid password.')
        # MDB=r"C:\Temp\#PrJRcd.mdb"  # 测试用临时目录
        connStr = f"DRIVER={{Microsoft Access Driver (*.mdb, *.accdb)}};DBQ={MDB};PWD=ABBabbELE;"
        try:
            self.conn = pypyodbc.connect(connStr, timeout=20, readonly=True)
            self.cursor = self.conn.cursor()
            return "success"
        except Exception as e:
            print("Cannot Connect to DB!!\n")
            print(e)
            return "fail"

    def readBOM(self,AVL):    
        # get the BOM from database
        sql_listTable = "SELECT BOM FROM DocAVL WHERE AVL='"+AVL+"' order by BOM"
        self.cursor.execute(sql_listTable)
        table_list = self.cursor.fetchall()
        #print(table_list) #测试输出的值
        return table_list
    
    def openMaxDB(self):
        connStr = "DSN=CIS_DESTO;Uid=LIMBAS2USER;Pwd=LIMBASREAD;"

        # 连接数据库
        try:
            self.conn = pypyodbc.connect(connStr, timeout=20, readonly=True)
            self.cursor = self.conn.cursor()
            print("Connect DB success!")
            return True
        except Exception as e:
            print("Cannot Connect to DB!!\n")
            return False
        

    def fetchMax(self, SAPNo_Searchby):
        # serach all table
        # fetch data
        sql_fetch = ''
        select_fields = 'PartNumber,SAP_Description,status,Detaildrawing,manufact_1,manufact_partnum_1,manufact_2,manufact_partnum_2,manufact_3,manufact_partnum_3,manufact_4,manufact_partnum_4,manufact_5,manufact_partnum_5,manufact_6,manufact_partnum_6,manufact_7,manufact_partnum_7'   #Different DB with different column name

        # SAP MAXDB检索区分大小写的COLLATE Latin1_General_CS_AS
        if SAPNo_Searchby != '':
            sql_append = "WHERE LOWER(SAP_Number) LIKE LOWER('{}')".format(SAPNo_Searchby)
        for index, tableName in enumerate(PartTypeList_CONNECT_4All_Search):
            # 每个table的SQL语句
            sql_each = "SELECT {} FROM {} ".format(select_fields, tableName)
            # SQL语句最后不添加;也不会出错的哦                        
            sql_each = sql_each + sql_append
            
            # 以下进行组合
            if index == 0:
                sql_fetch = sql_each
            else:                            
                sql_fetch = "{} UNION ALL ({})".format(sql_each,sql_fetch)
        
                     
        #print(sql_fetch+"\r\n")

        self.cursor.execute(sql_fetch)
        sql_result = self.cursor.fetchall()
        #print(sql_result)
        return sql_result

if __name__ == "__main__":
    try:
        # test the listTable function
        # print("Testing listTable function:")
        # db = Database()
        # for index, dbname in enumerate(DBList):
        #     print("=====================================")
        #     print("index:", index, "\ndbname:", dbname)
        #     db.openDB(index, DBList, None)
        #     db.listTable(index)
        
        # 
        db = Database()
        # 0: '01-CONNECT Local(ODBC)'; 1: '02-Access Online(ODBC)'; 2: '03-P Disk Access'; 3: '04-CONNECT DESTO(ODBC)'
        index = 0  
        db.openDB(index, DBList, None)
        # 生成PartNumber条件
        # PartNo_Searchby = "res_232"       
        PartNo_Searchby = ""       
         # 生成SAP_Number条件
        SAPNo_Searchby = "2tf"         
        # SAPNo_Searchby = ""          
        # 生成value条件
        PartValue_Searchby = "30K"    
        # PartValue_Searchby = "1U"    
        PartValue_Searchby = ""            
        # 生成manufact partnum 1-7的OR条件
        MfcPartNum_Searchby = "RC1206"   
        MfcPartNum_Searchby = ""   
        # 生成Description条件
        Description_Searchby = "0402"
        Description_Searchby = ""
        # 生成TechDescription条件
        TechDescription_Searchby = "FCN"
        # TechDescription_Searchby = ""
        # 生成Editor条件
        Editor_Searchby = "guozhaolin"
        Editor_Searchby = ""

        sql_result, columnNameList = db.fetch('---All----', index, PartNo_Searchby, SAPNo_Searchby, PartValue_Searchby, MfcPartNum_Searchby, Description_Searchby, TechDescription_Searchby, Editor_Searchby)
        print("Column Names:\n", columnNameList)
        print("=====================================")
        print("SQL Result:\n", sql_result)

    except Exception as e:
        print("Error:", e)
        pass





















