'''
todo:

'''


import pypyodbc


# windows parameter
# Database List can be used.
DBList = ['01-CONNECT Local(ODBC)', '02-Access Online(ODBC)', '03-P Disk Access','04-CONNECT DESTO(ODBC)']

# Part Type List for DB: '01-CONNECT Local(ODBC)'
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

# Part Type list for DB: '02-Access Online(ODBC)'
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
        # 01-CONNECT Local(ODBC)
        if dbindex == 0:
            connStr = "DSN=CIS_Local;Uid=LIMBAS2USER;Pwd=LIMBASREAD;"
            # connStr = "DSN=CONNECT Partslib V2;Uid=LIMBAS2USER;Pwd=LIMBASREAD;"
            print(dblist[dbindex])
        # 02-Access Online(ODBC)
        elif dbindex == 1:
            connStr = "DSN=CIS_PartLib_P_64_UNC;Uid=cadence_port;Pwd=Cadence_CIS.3;"
            print(dblist[dbindex])
        # 03-P Disk Access
        elif dbindex == 2:
            connStr = r"Driver={Microsoft Access Driver (*.mdb, *.accdb)};Dbq=P:\Cadence\CIS_DB_OL\CIS_PartLib.mdb;SystemDB=P:\Cadence\CIS_DB_OL\CIS_PartLib.mdw;Uid=cadence_port;Pwd=Cadence_CIS.3;"
            print(dblist[dbindex])
        elif dbindex == 3:
            connStr = "DSN=CIS_DESTO;Uid=LIMBAS2USER;Pwd=LIMBASREAD;"
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
        

    def fetch(self, tableName, dbindex, PartNo_Searchby, SAPNo_Searchby, PartValue_Searchby, MfcPartNum_Searchby):
        
        # if not search all
        if tableName != '---All----':
            # fetch data
            # 01-CONNECT Local(ODBC)
            if dbindex == 0 or dbindex == 3:
                # 无条件检索
                if (PartNo_Searchby == '') and (SAPNo_Searchby == '') and (PartValue_Searchby == '') and (MfcPartNum_Searchby == ''):
                    # 注意：SQL语句，最后不要添加;结束符号
                    sql_fetch = "SELECT * FROM {}".format(tableName)
                    # sql_fetch =  "SELECT * FROM RESISTORS where PARTNUMBER = 'RES_1868'"
                else:
                    print(PartNo_Searchby, SAPNo_Searchby, PartValue_Searchby, MfcPartNum_Searchby)
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
                    elif MfcPartNum_Searchby != '':
                        for index, MfcPartNum in enumerate(MftPartNumList_SAPMax):
                            if index == 0:
                                sql_append = "WHERE LOWER({}) LIKE LOWER(\'%{}%\')".format(MfcPartNum, MfcPartNum_Searchby)
                            else:
                                sql_append = "{} OR LOWER({}) LIKE LOWER(\'%{}%\')".format(sql_append, MfcPartNum, MfcPartNum_Searchby)

                    sql_fetch = sql_fetch + sql_append
                # SQL结果排序
                sql_fetch = sql_fetch + " ORDER BY PartNumber ASC" 
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
                if (PartNo_Searchby == '') and (SAPNo_Searchby == '') and (PartValue_Searchby == '') and (MfcPartNum_Searchby == ''):
                    sql_fetch = "SELECT * FROM [{}];".format(tableName)
                # 条件检索
                else:
                    print(PartNo_Searchby, SAPNo_Searchby, PartValue_Searchby, MfcPartNum_Searchby)
                    # 仅一个条件有效
                    sql_fetch = "SELECT * FROM [{}] ".format(tableName)
                    # SQL语句最后不添加;也不会出错的哦
                    if PartNo_Searchby != '':
                        sql_append = "WHERE PartNumber LIKE \'%{}%\'".format(PartNo_Searchby)
                    elif SAPNo_Searchby != '':
                        sql_append = "WHERE SAP_Number LIKE \'%{}%\'".format(SAPNo_Searchby)
                    elif PartValue_Searchby != '':
                        sql_append = "WHERE Value LIKE \'%{}%\'".format(PartValue_Searchby)
                    elif MfcPartNum_Searchby != '':
                        for index, MfcPartNum in enumerate(MftPartNumList_Access):
                            if index == 0:
                                sql_append = "WHERE {} LIKE \'%{}%\'".format(MfcPartNum, MfcPartNum_Searchby)
                            else:
                                sql_append = "{} OR {} LIKE \'%{}%\'".format(sql_append, MfcPartNum, MfcPartNum_Searchby)

                    sql_fetch = sql_fetch + sql_append
                # SQL结果排序
                sql_fetch = sql_fetch + " ORDER BY PartNumber ASC" 
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
            # 01-CONNECT Local(ODBC)
            if dbindex == 0 or dbindex == 3:
                select_fields = 'PartNumber,value_1,SAP_Number,SAP_Description,status,parttype,manufact_1,manufact_partnum_1,datasheet_1,manufact_2,manufact_partnum_2,datasheet_2,manufact_3,manufact_partnum_3,datasheet_3,manufact_4,manufact_partnum_4,datasheet_4,manufact_5,manufact_partnum_5,datasheet_5,manufact_6,manufact_partnum_6,datasheet_6,manufact_7,manufact_partnum_7,datasheet_7,scm_symbol,pcb_footprint,alt_symbols,mounttechn,ad_symbol,ad_footprint,ad_alt_footprint, detaildrawing'   #Different DB with different column name
                # 无条件检索
                if (PartNo_Searchby == '') and (SAPNo_Searchby == '') and (PartValue_Searchby == '') and (MfcPartNum_Searchby == ''):
                    # 注意：SQL语句，最后不要添加;结束符号
                    for index, tableName in enumerate(PartTypeList_CONNECT_4All_Search):
                        if index == 0:
                            sql_fetch = "SELECT {} FROM {}".format(select_fields, tableName)
                        else:
                            sql_fetch = "SELECT {} FROM {} UNION ALL ({})".format(select_fields, tableName,sql_fetch)
                    # print(sql_fetch)   
                # 条件检索
                else:
                    print(PartNo_Searchby, SAPNo_Searchby, PartValue_Searchby, MfcPartNum_Searchby)
                    # SAP MAXDB检索区分大小写的COLLATE Latin1_General_CS_AS
                    if PartNo_Searchby != '':
                        sql_append = "WHERE LOWER(PartNumber) LIKE LOWER(\'%{}%\')".format(PartNo_Searchby)
                    elif SAPNo_Searchby != '':
                        sql_append = "WHERE LOWER(SAP_Number) LIKE LOWER(\'%{}%\')".format(SAPNo_Searchby)
                    elif PartValue_Searchby != '':
                        sql_append = "WHERE LOWER(Value_1) LIKE LOWER(\'%{}%\')".format(PartValue_Searchby)
                    elif MfcPartNum_Searchby != '':
                        for index, MfcPartNum in enumerate(MftPartNumList_SAPMax):
                            if index == 0:
                                sql_append = "WHERE LOWER({}) LIKE LOWER(\'%{}%\')".format(MfcPartNum, MfcPartNum_Searchby)
                            else:
                                sql_append = "{} OR LOWER({}) LIKE LOWER(\'%{}%\')".format(sql_append, MfcPartNum, MfcPartNum_Searchby)
                    
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
                
                # SQL结果排序
                sql_fetch = sql_fetch + " ORDER BY PartNumber ASC"
                             
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
                select_fields = 'PartNumber,value,SAP_Number,SAP_Description,status,parttype,[manufact 1],[manufact partnum 1],[datasheet 1],[manufact 2],[manufact partnum 2],[datasheet 2],[manufact 3],[manufact partnum 3],[datasheet 3],[manufact 4],[manufact partnum 4],[datasheet 4],[manufact 5],[manufact partnum 5],[datasheet 5],[manufact 6],[manufact partnum 6],[datasheet 6],[manufact 7],[manufact partnum 7],[datasheet 7],scm_symbol,pcb_footprint,pcb_footprint_cp,alt_symbols,alt_symbols_cp,mounttechn,ad_symbol,ad_footprint,ad_alt_footprint,detaildrawing'   #Different DB with different column name
                # 无条件检索
                if (PartNo_Searchby == '') and (SAPNo_Searchby == '') and (PartValue_Searchby == '') and (MfcPartNum_Searchby == ''):
                    # 注意：SQL语句，最后不要添加;结束符号
                    for index, tableName in enumerate(PartTypeList_Access_4All_Search):
                        if index == 0:
                            sql_fetch = "SELECT {} FROM [{}]".format(select_fields, tableName)
                        else:
                            sql_fetch = "SELECT {} FROM [{}] UNION ALL ({})".format(select_fields, tableName,sql_fetch)
                    # print(sql_fetch)
                # 条件检索
                else:
                    print(PartNo_Searchby, SAPNo_Searchby, PartValue_Searchby, MfcPartNum_Searchby)
                    # SAP MAXDB检索区分大小写的COLLATE Latin1_General_CS_AS
                    if PartNo_Searchby != '':
                        sql_append = "WHERE PartNumber LIKE \'%{}%\'".format(PartNo_Searchby)
                    elif SAPNo_Searchby != '':
                        sql_append = "WHERE SAP_Number LIKE \'%{}%\'".format(SAPNo_Searchby)
                    elif PartValue_Searchby != '':
                        sql_append = "WHERE Value LIKE \'%{}%\'".format(PartValue_Searchby)
                    elif MfcPartNum_Searchby != '':
                        for index, MfcPartNum in enumerate(MftPartNumList_Access):
                            if index == 0:
                                sql_append = "WHERE {} LIKE \'%{}%\'".format(MfcPartNum, MfcPartNum_Searchby)
                            else:
                                sql_append = "{} OR {} LIKE \'%{}%\'".format(sql_append, MfcPartNum, MfcPartNum_Searchby)
                    
                    for index, tableName in enumerate(PartTypeList_Access_4All_Search):
                        # 每个table的SQL语句
                        sql_each = "SELECT {} FROM [{}] ".format(select_fields, tableName)
                        # SQL语句最后不添加;也不会出错的哦                        
                        sql_each = sql_each + sql_append
                        
                        # 以下进行组合
                        if index == 0:
                            sql_fetch = sql_each
                        else:                            
                            sql_fetch = "{} UNION ALL ({})".format(sql_each,sql_fetch)
                # SQL结果排序
                sql_fetch = sql_fetch + " ORDER BY PartNumber ASC"            
                print(sql_fetch)

                self.cursor.execute(sql_fetch)
                # columns = [column[0] for column in cursor.description]
                columnNameList = [column[0] for column in self.cursor.description]
                sql_result = self.cursor.fetchall()
                # print(sql_result)
                return sql_result, columnNameList
                        # 02-Access Online(ODBC) and 03-P Disk Access























