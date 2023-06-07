'''
todo:
1. 实现Part Type选择'---All----'时,数据库fetch读取的语句和功能
当前如果选择此项，则会报错
可以使用UNION ALL
-> CONNECT DB已经实现
next:
-> Access DB正在测试中，VPN网络太慢，需要等待回公司后再处理
-> UNION ALL的排序混乱，如何处理？
'''


import pypyodbc


# windows parameter
# Database List can be used.
DBList = ['01-CONNECT Online(ODBC)', '02-Access Online(ODBC)', '03-P Disk Access']

# Part Type List for DB: '01-CONNECT Online(ODBC)'
PartTypeList_CONNECT = [
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
('VARISTORS'),
('---All----')]


# Part Type list for DB: '02-Access Online(ODBC)'
PartTypeList_Access = [
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
 ('98-Shapes'),
 ('---All----')
]



# Database control class
class Database:
    def __ini__(self):
        # 初始化不需要创建任务东西
        pass
    
    def defaul(self,dbindex):
        # template
        # 01-CONNECT Online(ODBC)
        if dbindex == 0:
            pass
        # 02-Access Online(ODBC)
        elif dbindex == 1:
            pass
        # 03-P Disk Access
        elif dbindex == 2:
            pass
    
    def openDB(self, dbindex, dblist, app):
        # 01-CONNECT Online(ODBC)
        if dbindex == 0:
            connStr = "DSN=CONNECT Partslib V2;Uid=LIMBAS2USER;Pwd=LIMBASREAD;"
            print(dblist[dbindex])
        # 02-Access Online(ODBC)
        elif dbindex == 1:
            connStr = "DSN=CIS_PartLib_P_64;Uid=cadence_port;Pwd=Cadence_CIS.3;"
            print(dblist[dbindex])
        # 03-P Disk Access
        elif dbindex == 2:
            connStr = r"Driver={Microsoft Access Driver (*.mdb, *.accdb)};Dbq=P:\Cadence\CIS_DB_OL\CIS_PartLib.mdb;SystemDB=P:\Cadence\CIS_DB_OL\CIS_PartLib.mdw;Uid=cadence_port;Pwd=Cadence_CIS.3;"
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
    
    def listTable(self):    
        # get the table list
        sql_listTable = "SELECT NAME FROM MSYSOBJECTS WHERE TYPE=1 AND FLAGS=0;"
        self.cursor.execute(sql_listTable)
        table_list = self.cursor.fetchall()
        print(table_list)
        return table_list
        

    def fetch(self, tableName, dbindex, PartNo_Searchby, SAPNo_Searchby, PartValue_Searchby):
        
        # if not search all
        if tableName != '---All----':
            # fetch data
            # 01-CONNECT Online(ODBC)
            if dbindex == 0:
                # 无条件检索
                if (PartNo_Searchby == '') and (SAPNo_Searchby == '') and (PartValue_Searchby == ''):
                    # 注意：SQL语句，最后不要添加;结束符号
                    sql_fetch = "SELECT * FROM {}".format(tableName)
                    # sql_fetch =  "SELECT * FROM RESISTORS where PARTNUMBER = 'RES_1868'"
                else:
                    print(PartNo_Searchby, SAPNo_Searchby, PartValue_Searchby)
                    # 仅一个条件有效
                    sql_fetch = "SELECT * FROM {} ".format(tableName)
                    # SQL语句最后不添加;也不会出错的哦
                    # SAP MAXDB检索区分大小写的COLLATE Latin1_General_CS_AS
                    if PartNo_Searchby != '':
                        sql_append = "WHERE LOWER(PartNumber) LIKE LOWER(\'%{}%\')".format(PartNo_Searchby)
                    elif SAPNo_Searchby != '':
                        sql_append = "WHERE LOWER(SAP_Number) LIKE LOWER(\'%{}%\')".format(SAPNo_Searchby)
                    elif PartValue_Searchby != '':
                        sql_append = "WHERE LOWER(Value_1) LIKE LOWER(\'%{}%\')".format(PartValue_Searchby)
                    sql_fetch = sql_fetch + sql_append
                    print(sql_fetch)

                self.cursor.execute(sql_fetch)
                # columns = [column[0] for column in cursor.description]
                columnNameList = [column[0] for column in self.cursor.description]
                sql_result = self.cursor.fetchall()
                # print(sql_result)
                return sql_result, columnNameList
            # 02-Access Online(ODBC) and 03-P Disk Access
            elif dbindex == 1 or dbindex == 2:
                # 无条件检索
                if (PartNo_Searchby == '') and (SAPNo_Searchby == '') and (PartValue_Searchby == ''):
                    sql_fetch = "SELECT * FROM [{}];".format(tableName)
                # 条件检索
                else:
                    print(PartNo_Searchby, SAPNo_Searchby, PartValue_Searchby)
                    # 仅一个条件有效
                    sql_fetch = "SELECT * FROM [{}] ".format(tableName)
                    # SQL语句最后不添加;也不会出错的哦
                    if PartNo_Searchby != '':
                        sql_append = "WHERE PartNumber LIKE \'%{}%\'".format(PartNo_Searchby)
                    elif SAPNo_Searchby != '':
                        sql_append = "WHERE SAP_Number LIKE \'%{}%\'".format(SAPNo_Searchby)
                    elif PartValue_Searchby != '':
                        sql_append = "WHERE Value LIKE \'%{}%\'".format(PartValue_Searchby)
                    sql_fetch = sql_fetch + sql_append
                    print(sql_fetch)

                self.cursor.execute(sql_fetch)
                # columns = [column[0] for column in cursor.description]
                columnNameList = [column[0] for column in self.cursor.description]
                sql_result = self.cursor.fetchall()
                # print(sql_result)
                return sql_result, columnNameList
            # 03-P Disk Access
            # elif dbindex == 2:
            #     pass
        # serach all table
        else:
            # fetch data
            # select_fields = 'PartNumber,value,SAP_Number,SAP_Description'
            
            sql_fetch = ''
            # 01-CONNECT Online(ODBC)
            if dbindex == 0:
                select_fields = 'PartNumber,value_1,SAP_Number,SAP_Description'   #Different DB with different column name
                # 无条件检索
                if (PartNo_Searchby == '') and (SAPNo_Searchby == '') and (PartValue_Searchby == ''):
                    # 注意：SQL语句，最后不要添加;结束符号
                    for index, tableName in enumerate(PartTypeList_CONNECT[:-1]):
                        if index == 0:
                            sql_fetch = "SELECT {} FROM {}".format(select_fields, tableName)
                        else:
                            sql_fetch = "SELECT {} FROM {} UNION ALL ({})".format(select_fields, tableName,sql_fetch)
                    # print(sql_fetch)
                # 条件检索
                else:
                    print(PartNo_Searchby, SAPNo_Searchby, PartValue_Searchby)
                    # SAP MAXDB检索区分大小写的COLLATE Latin1_General_CS_AS
                    if PartNo_Searchby != '':
                        sql_append = "WHERE LOWER(PartNumber) LIKE LOWER(\'%{}%\')".format(PartNo_Searchby)
                    elif SAPNo_Searchby != '':
                        sql_append = "WHERE LOWER(SAP_Number) LIKE LOWER(\'%{}%\')".format(SAPNo_Searchby)
                    elif PartValue_Searchby != '':
                        sql_append = "WHERE LOWER(Value_1) LIKE LOWER(\'%{}%\')".format(PartValue_Searchby)
                    for index, tableName in enumerate(PartTypeList_CONNECT[:-1]):
                        # 每个table的SQL语句
                        sql_each = "SELECT {} FROM {} ".format(select_fields, tableName)
                        # SQL语句最后不添加;也不会出错的哦                        
                        sql_each = sql_each + sql_append
                        
                        # 以下进行组合
                        if index == 0:
                            sql_fetch = sql_each
                        else:                            
                            sql_fetch = "{} UNION ALL ({})".format(sql_each,sql_fetch)
                            
                print(sql_fetch)

                self.cursor.execute(sql_fetch)
                # columns = [column[0] for column in cursor.description]
                columnNameList = [column[0] for column in self.cursor.description]
                sql_result = self.cursor.fetchall()
                # print(sql_result)
                return sql_result, columnNameList
                        # 02-Access Online(ODBC) and 03-P Disk Access
            # 02-Access Online(ODBC) and 03-P Disk Access
            elif dbindex == 1 or dbindex == 2: 
                select_fields = 'PartNumber,value,SAP_Number,SAP_Description'   #Different DB with different column name
                # 无条件检索
                if (PartNo_Searchby == '') and (SAPNo_Searchby == '') and (PartValue_Searchby == ''):
                    # 注意：SQL语句，最后不要添加;结束符号
                    for index, tableName in enumerate(PartTypeList_Access[:-1]):
                        if index == 0:
                            sql_fetch = "SELECT {} FROM [{}]".format(select_fields, tableName)
                        else:
                            sql_fetch = "SELECT {} FROM [{}] UNION ALL ({})".format(select_fields, tableName,sql_fetch)
                    # print(sql_fetch)
                # 条件检索
                else:
                    print(PartNo_Searchby, SAPNo_Searchby, PartValue_Searchby)
                    # SAP MAXDB检索区分大小写的COLLATE Latin1_General_CS_AS
                    if PartNo_Searchby != '':
                        sql_append = "WHERE PartNumber LIKE \'%{}%\'".format(PartNo_Searchby)
                    elif SAPNo_Searchby != '':
                        sql_append = "WHERE SAP_Number LIKE \'%{}%\'".format(SAPNo_Searchby)
                    elif PartValue_Searchby != '':
                        sql_append = "WHERE Value LIKE \'%{}%\'".format(PartValue_Searchby)
                    for index, tableName in enumerate(PartTypeList_Access[:-1]):
                        # 每个table的SQL语句
                        sql_each = "SELECT {} FROM [{}] ".format(select_fields, tableName)
                        # SQL语句最后不添加;也不会出错的哦                        
                        sql_each = sql_each + sql_append
                        
                        # 以下进行组合
                        if index == 0:
                            sql_fetch = sql_each
                        else:                            
                            sql_fetch = "{} UNION ALL ({})".format(sql_each,sql_fetch)
                            
                print(sql_fetch)

                self.cursor.execute(sql_fetch)
                # columns = [column[0] for column in cursor.description]
                columnNameList = [column[0] for column in self.cursor.description]
                sql_result = self.cursor.fetchall()
                # print(sql_result)
                return sql_result, columnNameList
                        # 02-Access Online(ODBC) and 03-P Disk Access























