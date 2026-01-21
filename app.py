'''
LastEditors: Mackey
LastEditTime: 2024-04-29 15:30:00
Introduction: 
    Flask网页端访问CONNECT DB并显示
    即制作网页版的CONNECT DB显示程序

Revision History:
1.0.0 - 20240422: 初始版本,实现CONNECT DB的网页端查询和Excel保存功能
1.0.1 - 20240429: 修正了Excel保存功能的bug,之前保存的Excel文件无法打开
1.1.0 - 20240506: 增加了对Access DB的支持,可以选择CONNECT DB和Access DB进行查询
1.1.1 - 20240510: 修正了Access DB查询时的bug,之前查询结果不正确
1.2.0 - 20240515: 优化了网页界面,增加了查询条件的输入框
1.2.1 - 20240520: 修正了网页界面的一些显示问题,提升用户体验
1.3.0 - 20240525: 增加了对AVL BOM导出的支持,可以从Windchill获取BOM并生成Excel文件, by Cyrus
2.0.0 - 20260108: 重构代码结构,db_mgt.py使用PyPika进行SQL语句生成,提升代码可维护性和扩展性
2.1.0 - 20260108: 将搜索的条件显示在页面
2.2.0 - 20260108: 搜索条件输入内容在点击search之后不会清除
2.3.0 - 20260108: 新增过滤条件“SAP_Description”"techdescription" "editor", 网页端增加输入框。
2.4.0 - 20260108: 支持多个SAP编号的批量查询,输入多个SAP编号,按空格、逗号、分号分隔, 未完成保存Excel功能
2.5.0 - 20260111: 完成多个SAP编号的批量查询结果直接保存为Excel文件功能
2.6.0 - 20260111: 在网页端显示软件版本号
3.0.0 - 20260118: 增加了AVL处理页面,支持从Windchill获取BOM,查询数据库,生成AVL Excel文件,并通过AJAX方式下载文件
        a) 网站打开首页增加跳转到AVL处理页面的按钮, 并添加版本号显示
        b) 新增AVL处理页面,支持Create AVL和Download AVL功能,使用AJAX方式处理请求和下载文件
        c) AVL处理页面中的AVL inlcude选项支持“2TFU CN only”和“All”,默认为“2TFU CN only”, 但All选项还存在问题,需要后续修正
3.1.0 - 20260118: AVL页面添加跳转回主页面的按钮
3.1.1 - 20260119: 修正了AVL处理页面中的bug: AVL_include选项为"All Parts"时,未正确输出找不到ordering information的Parts导出Excel文件的问题。
3.2.0 - 20260119: 优化了AVL处理页面的问题, 如果输入Windchill用户名和密码为空, PCBA part number为空, 则提示并不继续处理
3.3.0 - 20260119: 增加判断是否获取到ordering information, 若没有则不继续处理,并提示用户
3.3.1 - 20260119: 优化PLM登录失败的处理逻辑,避免后续函数调用出错。
            PLM_Basic_Auth_ByPass_MFA_Get_BOM.py升级到1.4.0版本,get_BOM()函数增加返回PLM登录是否成功的标志PLM_Login_OK。
3.3.2 - 20260119: 修正CONNECT Viewer页面中的SAP Number List Search功能的bug
            当SAP编号未找到时,会进行判断，并添加空行占位,填写SAP number
3.4.0 - 20260119: "Download_AVL"按键实现下载功能
3.5.0 - 20260121: 增加AVL Comparison功能,支持上传手动整理的AVL文件进行对比,并生成对比结果Excel文件供下载。此功能暂时不支持自动生成AVL_Cmp sheet,需要用户手动整理后上传进行对比。临时版本号提升为3.5.0,等待后续完善自动生成AVL_Cmp sheet功能。
3.6.0 - 20260121: 增加Compare_Manual_AVL按键, 用于上传手动整理的AVL文件进行对比,并生成对比结果Excel文件供下载。
'''

# 版本号
# xx.yy.zz
# xx: 大版本，架构性变化
# yy: 功能性新增
# zz: Bug修复
__Version__ = "3.6.0"

import sys
from flask import Flask, send_file , jsonify , request, redirect
from flask import render_template
from flask.helpers import flash, url_for
import pandas as pd  
import openpyxl
import io  
import db_mgt
import re  
import requests  
import base64 
import json  
import logging
import datetime
from werkzeug.utils import secure_filename
import os
import third_party.PLM_Handle.PLM_Basic_Auth_ByPass_MFA_Get_BOM as plm
import third_party.Excel_Handle.AVL_Excel_Handle as excel_handle
import openpyxl
import tempfile
import os

# Global Variables
first_AVL_Output_File = ""  # Excel输出文件路径
AVL_Compare_Output_File = ""  # AVL比较输出文件路径

# debug print, print到控制台
DEBUG_PRINT = True

def debug_print(*args, **kwargs):
    if DEBUG_PRINT:
        print(*args, **kwargs)

# production logging
# from logging.config import dictConfig

# dictConfig({
#     'version': 1,
#     'formatters': {'default': {
#         'format': '[%(asctime)s] %(levelname)s in %(module)s: %(message)s',
#     }},
#     'handlers': {
#         'wsgi': {
#             'class': 'logging.StreamHandler',
#             'formatter': 'default'
#         },
#         'custom_handler': {
#             'class': 'logging.FileHandler',
#             'formatter': 'default',
#             'filename': r'D:\80_MackeyDoc\01_ABB\OneDrive - ABB\00_WorkPlace\01_Design_Work\01_Prg\10_Py\PythonCode\A1_Flask\Flask_CONNECT_DB\myapp.log'
#         }
#     },
#     'root': {
#         'level': 'DEBUG',
#         'handlers': ['wsgi', 'custom_handler']
#     }
# })



app = Flask(__name__)

# 对于IIS生产系统,这段要放在这里,不能放在main里
# create DB instance
db = db_mgt.Database()

# set debug, can work even in production
app.config['ENV'] = 'development'
app.config['DEBUG'] = True
app.config['TESTING'] = True

app.secret_key = 'CONNECT'

# Set the loggin
# handler = logging.FileHandler('app.log')  # errors logged to this file
# handler.setLevel(logging.NOTSET)  # only log errors and above
# app.logger.addHandler(handler)  # attach the handler to the app's logger
# app.logger.info("App start !")

# create DB instance
WC_Path = "https://lp-global-plm.abb.com/Windchill/protocolAuth/servlet/odata/"
#定义一个空的集合用于记录AVL里的元器件清单
AVLPart_ListView=set()
#定义一个空的集合用于在网页端显示内容以供使用者检查
Component_ListView=set()
#定义Excel模板中,有效数据的首行
Excel_Row=7


'''
初始始页面： 选择数据库
点击确认之后,会跳转到检索页面,并且会将选择的数据库序号传递过去。
'''
@app.route("/", methods=['GET','POST'])
def DBSelect():
    DB_List=db_mgt.DBList   
    
    if request.method == "POST":
        # 获取选择的数据库类型, 为str类型
        DB_Select = request.form.get("DB_Select")
        DB_Index = db_mgt.DBList.index(DB_Select)
        print (DB_Index)
        return redirect(url_for('index',DBType=DB_Index))
    else:
        return render_template('DBSelect.html', DB_List=DB_List,Version = __Version__)

'''
数据库检索和显示页面：初始始显示检索表单,在此之前会尝试打开数据库
submit之后显示检索内容
'''
@app.route("/search/<DBType>", methods=['GET','POST'])
def index(DBType):
    # 根据DBType来设置Part Type 列表的内容,DBType为str,对应db_mgt.DBList的index值,从0开始
    if DBType == '0' or DBType == '3': 
        #如果将值直接在render_template里赋值,数据第一次会传递不过去,不知原因。
        Part_Type_List=db_mgt.PartTypeList_CONNECT
    elif DBType == '1':
        Part_Type_List=db_mgt.PartTypeList_Access
    else:
        Part_Type_List=db_mgt.PartTypeList_Access
    
    # 尝试打开DB
    # Open DB
    bIsDBOpen = db.openDB(int(DBType), db_mgt.DBList, app)
    if bIsDBOpen == True:
        flash(db_mgt.DBList[int(DBType)]+" 打开数据库成功！")
    else:
        flash(db_mgt.DBList[int(DBType)]+" 打开数据库出错")
    
    # 提交表单查询时处理
    if request.method == "POST":
        # 定义全局变量
        global sql_result
        global columnNameList
        global MaxLine
        global sql_result_len
        dbindex = int(DBType)      
        # 获取表单内容
        tableName  = request.form.get("tableName")        
        PartNo_Searchby = request.form.get("PartNo")
        SAPNo_Searchby = request.form.get("SAPNo")
        PartValue_Searchby = request.form.get("PartValue")
        MfcPartNum_Searchby = request.form.get("MfcPartNum")
        MaxLine_str = request.form.get("MaxLine")
        MaxLine = int(MaxLine_str) if MaxLine_str is not None and MaxLine_str.strip() != "" else 100  # default to 100 if not provided
        Description_Searchby = request.form.get("Description")
        TechDescription_Searchby = request.form.get("TechDescription")
        Editor_Searchby = request.form.get("Editor")        
        SAP_Number_List = request.form.get("SAP_Number_List")
            
                  
        # 判断按键
        print("press the button:" + request.form['btn'])
        if request.form['btn'] == 'Search':
            # 条件搜索
            Search_Info = f"PartNo: {PartNo_Searchby}, SAPNo: {SAPNo_Searchby}, PartValue: {PartValue_Searchby}, MfcPartNum: {MfcPartNum_Searchby}, MaxLine: {MaxLine}, TableName: {tableName}"
            # 执行搜索
            sql_result, columnNameList = db.fetch(
                tableName=tableName, 
                dbindex=dbindex, 
                PartNo_Searchby=PartNo_Searchby, 
                SAPNo_Searchby=SAPNo_Searchby, 
                PartValue_Searchby=PartValue_Searchby, 
                MfcPartNum_Searchby=MfcPartNum_Searchby, 
                Description_Searchby=Description_Searchby, 
                TechDescription_Searchby=TechDescription_Searchby, 
                Editor_Searchby=Editor_Searchby
            )
            sql_result_len = len(sql_result)
            # 显示结果
            return render_template(
                'index.html', 
                Part_Type_List=Part_Type_List, 
                MaxLine=MaxLine, 
                sql_result=sql_result, 
                columnNameList=columnNameList, 
                sql_result_len=sql_result_len, 
                Search_Info=Search_Info,
                PartNo=PartNo_Searchby,
                SAPNo=SAPNo_Searchby,
                PartValue=PartValue_Searchby,
                MfcPartNum=MfcPartNum_Searchby,
                tableName=tableName,
                Description=Description_Searchby,
                TechDescription=TechDescription_Searchby,
                Editor=Editor_Searchby,
                Version=__Version__
                )
            # return db_mgt.DBList[0]
        elif request.form['btn'] == 'SaveExcel':
            # 保存Excel
            print("SaveExcel")
            # print(columnNameList)
            if 'columnNameList' in globals():
                temp_dir = tempfile.gettempdir()    # not used, as in the server the temp dir is not the same as in the local
                # file_path = os.path.join(os.path.dirname(os.path.realpath(sys.argv[0])), 'ExportFiles', "SQL_Result.xlsx")
                # import datetime
                # 用唯一文件名（如加时间戳）,避免冲突
                filename = f"SQL_Result_{datetime.datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx"
                file_path = os.path.join(os.path.dirname(os.path.realpath(sys.argv[0])), 'ExportFiles', filename)
                if file_path:
                    wb = openpyxl.Workbook()
                    ws = wb.active
                    ws.append(columnNameList)
                    for row in sql_result:
                        ws.append(row)
                    wb.save(file_path)
                    flash("Excel保存成功！{}".format(file_path))
                    # 打开Excel, 文件会保存在服务器中,客户端是无法直接打开这个文件的,此方法行不通的。
                    # print(file_path)                    
                    # os.system('start excel.exe {}'.format('"' + file_path + '"'))
                # 可以使用send_file来发送文件给客户端
                flash("Excel保存成功！")
                return send_file(file_path, as_attachment=True)
                # return render_template('index.html', Part_Type_List=Part_Type_List, MaxLine=MaxLine, sql_result=sql_result, columnNameList=columnNameList, sql_result_len=sql_result_len)
            else:
                flash("没有数据,无法保存Excel！")
                return render_template('index.html', Part_Type_List=Part_Type_List,Version=__Version__)
        elif request.form['btn'] == 'SAP_Nums_Search':
            # 搜索多个SAP编号
            # 处理多个SAP编号的输入,按空格、逗号、分号分隔
            print("SAP_Nums_Search")
            SAP_Nums_List = re.split(r'[\s,;]+', SAP_Number_List.strip())
            print(SAP_Nums_List)

            sql_result = []
            for SAPNo_Searchby in SAP_Nums_List:
                sql_result_each, columnNameList = db.fetch(
                    tableName=tableName, 
                    dbindex=dbindex, 
                    PartNo_Searchby='',
                    SAPNo_Searchby=SAPNo_Searchby,
                    PartValue_Searchby='',
                    MfcPartNum_Searchby='',
                    Description_Searchby='',
                    TechDescription_Searchby='',
                    Editor_Searchby=''
                    )
                if sql_result_each: # 非空结果才添加
                    sql_result.append(sql_result_each[0])
                else: # 未找到结果, 添加空行占位,填写SAP number
                    sql_result.append(['']*2 + [SAPNo_Searchby] + [''] * (len(columnNameList) - 3))
            sql_result_len = len(sql_result)
            return render_template(
                'index.html',
                Part_Type_List=Part_Type_List,
                sql_result=sql_result,
                columnNameList=columnNameList,  
                sql_result_len=sql_result_len,
                SAP_Nums=SAP_Number_List,
                MaxLine=MaxLine,
                Version=__Version__
                )
        else:
            return render_template('index.html', Part_Type_List=Part_Type_List, Version=__Version__)
    else:
        return render_template('index.html', Part_Type_List=Part_Type_List, Version=__Version__)

# =========== 以下部分为2026/01 之后新的AVL处理代码 ==============
def download_excel(output_excel_file, AJAX=False, msg_avlHandle="", btn_enabled=True):
    """ 弹窗下载AVL Excel文件。
    Args:
        param output_excel_file (str): 输出Excel文件路径
    return: 
        output_excel_file (str): 输出Excel文件路径
    """
    # 这里可以添加下载文件的逻辑
    if AJAX:
        # 生成Excel后
        download_url = url_for('downloadExcelFile', filename=os.path.basename(output_excel_file))
        return jsonify({'status': 'completed', 
                        'msg': msg_avlHandle, 
                        'btn_enabled': btn_enabled, 
                        'download_url': download_url
                        })
    else:
        response_file = send_file(output_excel_file, as_attachment=True)
        return response_file

def get_ordering_info_from_db(db, dbindex, tableName, SAP_Number_List, Multi_BOM_Info_list):
    """ 根据SAP编号列表,查询数据库,返回sql_result和columnNameList
    Args:
        db(db_mgt.Database): 数据库实例
        dbindex(int): 数据库索引
        tableName(str): 表名
        SAP_Number_List(list): SAP编号列表
        Multi_BOM_Info_list(list): BOM信息列表,用于查不到信息时, 获取SAP Description以便填写

    Returns:
        tuple: (sql_result, columnNameList): 查询结果和列名列表
    """
    sql_result = []
    columnNameList = None
    for SAPNo_Searchby in SAP_Number_List:
        sql_result_each, columnNameList = db.fetch(
            tableName=tableName, 
            dbindex=dbindex, 
            PartNo_Searchby='',
            SAPNo_Searchby=SAPNo_Searchby,
            PartValue_Searchby='',
            MfcPartNum_Searchby='',
            Description_Searchby='',
            TechDescription_Searchby='',
            Editor_Searchby=''
        )
        if sql_result_each:
            sql_result.append(sql_result_each[0])
        else:
            # 未找到结果, 添加空行占位,填写SAP number和Description
            SAP_Description_Searchby = ""
            for bom_info in Multi_BOM_Info_list:
                if bom_info.split(',')[0] == SAPNo_Searchby:
                    SAP_Description_Searchby = bom_info.split(',')[1]
                    break
            Not_Found_SAP_info = ('','', SAPNo_Searchby, SAP_Description_Searchby)
            sql_result.append(Not_Found_SAP_info)
    return sql_result, columnNameList

# 通过URL跳转的方式下载Excel文件
@app.route('/downloadExcelFile/<filename>')
def downloadExcelFile(filename):
    filePathName = os.path.join(os.path.dirname(os.path.realpath(sys.argv[0])), 'ExportFiles', filename)
    return send_file(filePathName, as_attachment=True)

@app.route("/AVLhandle", methods=['GET','POST'])
def AVLHandle():
    # test flash('msg test'), 需要进行手动清除
    # flash('Welcome to AVL Handle Page!')
    # 网页运行信息
    msg_avlHandle = ""
    # 按键使能状态
    btn_enabled = True
    # 处理POST请求, 即点击按钮之后,判断按键类型并处理
    
    if request.method == 'POST':
        # 处理POST请求
        # 定义全局变量
        global first_AVL_Output_File  # 声明为全局变量, 第一次生成的AVL Excel文件路径
        global AVL_Compare_Output_File  # 声明为全局变量, 对比后的AVL Excel文件路径
        # 获取public部分设置
        user = request.form.get('user')
        pwd = request.form.get('password')
        DB_Select = request.form.get('DB_Select')
        AVL_include = request.form.get("AVL_include")
        # 获取First AVL Generation部分设置
        PCBA_Part_Number_List = request.form.get('PCBA_Part_Number_List')
        btn = request.form.get('btn')
        # 获取AVL Comparison部分设置
        AVL_Cmp_range = request.form.get("AVL_Cmp_range")
        excel_file = request.files.get('excel_file')

        # debug print
        debug_print("user:", user)
        debug_print("pwd:", pwd)
        debug_print("DB_Select:", DB_Select)
        debug_print("AVL_include:", AVL_include)
        debug_print("PCBA_Part_Number_List:", PCBA_Part_Number_List)
        debug_print("btn:", btn)
        debug_print("AVL_Cmp_range:", AVL_Cmp_range)
        debug_print("excel_file:", excel_file)

        # 判断是否输入账号密码
        if (not user or user.strip() == "") or (not pwd or pwd.strip() == ""):
            msg_avlHandle = "Windchill user name and password cannot be empty. Please input valid credentials."
            return jsonify({
                'status': 'error', 
                'msg': msg_avlHandle,
                'btn_enabled': btn_enabled
            })


        # 以下代码用于处理按钮点击后,界面显示和按钮使能状态,已经在页面的JavaScript中实现,这里注释掉
        # disable all buttons during processing
        # btn_enabled = False  
        # msg_avlHandle = "正在处理,请稍候..."
        
        # 处理Create AVL按钮点击事件
        if btn == 'Create_AVL':
            debug_print("="*30)
            debug_print("Create_AVL button clicked. Start processing...")
            # step1: 准备工作,处理输入参数
            # 判断PCBA Part Number List是否为空
            if not PCBA_Part_Number_List or PCBA_Part_Number_List.strip() == "":
                msg_avlHandle = "PCBA Part Number list cannot be empty. Please input valid PCBA Part Numbers."
                return jsonify({
                    'status': 'error', 
                    'msg': msg_avlHandle,
                    'btn_enabled': btn_enabled
                })
            # 处理 PCBA_Part_Number_List,获取list格式
            PCBA_Part_Number_List = re.split(r'[\s,;]+', PCBA_Part_Number_List.strip())
            # 处理AVL_include选项
            bCHINA_PN_ONLY = True if AVL_include == '2TFU CN only' else False
            # 变量定义
            Multi_BOM_Info_list = []    # PartNumber,PartName,Quantity,DesignatorRange
            Multi_BOM_SAP_Number_List = []  # SAP Number list in the BOM
            Multi_BOM_SAP_Number_List_Str = ""  # SAP Number list in the BOM, str format
            Multi_PCBA_Part_info_list = []  # PCBA Part info list
            # Step2: 连接PLM,获取BOM信息
            debug_print("Starting to get BOM from PLM...")
            for BOM_number in PCBA_Part_Number_List:
                print("Processing PCBA Part Number:", BOM_number)
                BOM_Info_list, BOM_SAP_Number_List, BOM_SAP_Number_List_Str, PCBA_Part_info_list, PLM_Login_OK = plm.get_BOM(user, pwd, BOM_number, bCHINA_PN_ONLY)
                if not PLM_Login_OK:
                    debug_print(f"Failed to login to Windchill for PCBA Part Number: {BOM_number}")
                    msg_avlHandle = "Failed to login to Windchill. Please check your username, password and network connection."
                    return jsonify({
                        'status': 'error', 
                        'msg': msg_avlHandle,
                        'btn_enabled': btn_enabled
                    })
                else:
                    debug_print(f"BOM Info for {BOM_number} retrieved successfully.")
                    Multi_BOM_Info_list += BOM_Info_list
                    Multi_BOM_SAP_Number_List += BOM_SAP_Number_List
                    Multi_BOM_SAP_Number_List_Str += BOM_SAP_Number_List_Str
                    Multi_PCBA_Part_info_list.append(PCBA_Part_info_list)
            # 去除重复项
            Multi_BOM_SAP_Number_List = list(set(Multi_BOM_SAP_Number_List))
            Multi_BOM_Info_list = list(set(Multi_BOM_Info_list))
            # Step3: 通过SQL获取ordering information
            debug_print("Starting to get ordering info from DB...")
            tableName = '---All----' #检索所有表
            dbindex = int(DB_Select)    #0: CONNECT DB, 1: Access DB
            # 打开DB
            bIsDBOpen = db.openDB(dbindex, db_mgt.DBList, app)
            if bIsDBOpen == True:
                flash(db_mgt.DBList[dbindex]+" 打开数据库成功！")
            else:
                flash(db_mgt.DBList[dbindex]+" 打开数据库出错")
                return jsonify({
                    'status': 'error',
                    'msg': "Failed to open the selected database.",
                    'btn_enabled': btn_enabled
                })
            # 调用重构后的函数，获取ordering information
            sql_result, columnNameList = get_ordering_info_from_db(
                db, dbindex, tableName, Multi_BOM_SAP_Number_List, Multi_BOM_Info_list
            )
            # 判断是否有查询到数据，若没有则不继续处理，并提示用户
            sql_result_len = len(sql_result)
            if sql_result_len == 0:
                msg_avlHandle = "No ordering information found for the given PCBA Part Numbers' BOM SAP Numbers."
                return jsonify({
                    'status': 'error', 
                    'msg': msg_avlHandle,
                    'btn_enabled': btn_enabled
                })
            # Step4: 保存Excel文件, 使用模板2TFP900033A1076.xlsx
            debug_print("Starting to write AVL Excel file...")
            # Excel模板路径
            template_file = os.path.join(os.path.dirname(os.path.realpath(sys.argv[0])),  '2TFP900033A1076.xlsx')
            # 输出文件路径
            first_AVL_Output_File = os.path.join(os.path.dirname(os.path.realpath(sys.argv[0])), 'ExportFiles', f"AVL_Result_{datetime.datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx")
            # 调用Excel处理模块,生成AVL Excel文件, 结果直接输出到output_file
            excel_handle.first_write_AVL_to_excel(template_file, sql_result, Multi_PCBA_Part_info_list, first_AVL_Output_File)
            # Step5: 完成操作,返回结果信息
            debug_print("AVL Excel file created successfully.")
            # Step6: 提供下载
            msg_avlHandle = "Create AVL button processing completed. If the save dialog did not pop up, please click the Download_AVL button."
            # AJAX方式下载文件
            return download_excel(first_AVL_Output_File, AJAX=True, msg_avlHandle=msg_avlHandle, btn_enabled=True)

        # 处理Download AVL按钮点击事件
        elif btn == 'Download_AVL':
            debug_print("="*30)
            debug_print("Download_AVL button clicked. Start processing...")
            if not first_AVL_Output_File or not os.path.exists(first_AVL_Output_File):
                msg_avlHandle = "AVL Excel file not found. Please click the Create AVL button first."
                return jsonify({
                    'status': 'error', 
                    'msg': msg_avlHandle,
                    'btn_enabled': btn_enabled
                })
            # AJAX方式下载文件
            debug_print("Download AVL successfully.")
            msg_avlHandle = "Download_AVL button processing completed."
            return download_excel(first_AVL_Output_File, AJAX=True, msg_avlHandle=msg_avlHandle, btn_enabled=True)
        # 处理Manual Compare AVL按钮点击事件
        elif btn == 'Compare_Manual_AVL':
            debug_print("="*30)
            debug_print("Compare_Manual_AVL button clicked. Start processing...")
            # step1: 准备工作,处理输入参数
            # 判断excel_file是否为空
            if not excel_file:
                msg_avlHandle = "Please upload an Excel file for AVL comparison."
                return jsonify({
                    'status': 'error', 
                    'msg': msg_avlHandle,
                    'btn_enabled': btn_enabled
                })
            # 保存上传的Excel文件到临时路径
            temp_dir = tempfile.gettempdir()
            uploaded_file = os.path.join(temp_dir, secure_filename(excel_file.filename))
            excel_file.save(uploaded_file)
            debug_print("Uploaded Excel file saved to:", uploaded_file)
            # 判断Excel文件中是否包含"AVL"和"AVL_Cmp"两个sheet
            required_sheets = excel_handle.AVL_MANUAL_REQUIRED_SHEETS
            if not excel_handle.check_AVL_file(uploaded_file, required_sheets):
                msg_avlHandle = f"The uploaded Excel file must contain the following sheets: {', '.join(required_sheets)}."
                return jsonify({
                    'status': 'error', 
                    'msg': msg_avlHandle,
                    'btn_enabled': btn_enabled
                })
            # 输出文件路径 
            AVL_Compare_Output_File = os.path.join(os.path.dirname(os.path.realpath(sys.argv[0])), 'ExportFiles', f"AVL_Compare_Result_{datetime.datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx")
            excel_handle.compare_avl_sheets(uploaded_file, AVL_Compare_Output_File)
            msg_avlHandle = "Compare Manual AVL button processing completed. If the save dialog did not pop up, please click the Download_Result button."
            # AJAX方式下载文件
            return download_excel(AVL_Compare_Output_File, AJAX=True, msg_avlHandle=msg_avlHandle, btn_enabled=True)
        # 处理Compare AVL按钮点击事件
        elif btn == 'Compare_AVL':
            debug_print("="*30)
            debug_print("Compare_AVL button clicked. Start processing...")
            # step1: 准备工作,处理输入参数
            # 判断excel_file是否为空
            if not excel_file:
                msg_avlHandle = "Please upload an Excel file for AVL comparison."
                return jsonify({
                    'status': 'error', 
                    'msg': msg_avlHandle,
                    'btn_enabled': btn_enabled
                })
            # 保存上传的Excel文件到临时路径
            temp_dir = tempfile.gettempdir()
            uploaded_file = os.path.join(temp_dir, secure_filename(excel_file.filename))
            excel_file.save(uploaded_file)
            debug_print("Uploaded Excel file saved to:", uploaded_file)
            # 判断Excel文件中是否包含"AVL" sheet
            required_sheets = excel_handle.AVL_AUTO_REQUIRED_SHEETS
            if not excel_handle.check_AVL_file(uploaded_file, required_sheets):
                msg_avlHandle = f"The uploaded Excel file must contain the following sheet: {', '.join(required_sheets)}."
                return jsonify({
                    'status': 'error', 
                    'msg': msg_avlHandle,
                    'btn_enabled': btn_enabled
                })
            # step2: 根据Compare范围, 连接数据库生成AVL_Cmp sheet
            debug_print("Starting to create AVL_Cmp sheet...")
            # Excel模板路径
            template_file = os.path.join(os.path.dirname(os.path.realpath(sys.argv[0])),  '2TFP900033A1076.xlsx')
            # todo: 1. 根据AVL_Cmp_range选择不同的处理方式, 生成AVL_Cmp sheet
            # 选项1: AVL Sheet Only, 获取上传文件中的AVL sheet内容，然后根据B列第7行开始的Part Numbers查询数据库，生成AVL_Cmp sheet
            if AVL_Cmp_range == 'AVL_Sheet_Only':
                SAP_Nums_List = excel_handle.get_SAP_Numbers_from_AVL_sheet(uploaded_file)
                debug_print("SAP Numbers extracted from AVL sheet:", SAP_Nums_List)
                debug_print("len(SAP_Nums_List):", len(SAP_Nums_List))
                pass
            # 选项2: BOM Related sheet,  获取上传文件中的BOM Related sheet内容，然后根据B列第3行开始的PCBA Part Numbers查询PLM获取BOM，然后根据BOM中的SAP Numbers查询数据库，生成AVL_Cmp sheet
            elif AVL_Cmp_range == 'BOM_Related_Sheet':
                # 此部分的功能与Create AVL部分类似, 需要重构代码以便复用
                pass

            # 
            # 调用Excel处理模块,生成AVL Comparison Excel文件, 结果直接输出到output_file



            debug_print("Compare AVL successfully.")
            msg_avlHandle = "Create AVL button processing completed. If the save dialog did not pop up, please click the Download_Result button."
        # 处理Download Result按钮点击事件
        elif btn == 'Download_Result':
            debug_print("="*30)
            debug_print("Download_Result button clicked. Start processing...")

            debug_print("Download Result successfully.")
            msg_avlHandle = "Download_Result button processing completed."
        # 处理未知按钮点击事件
        else:
            debug_print("="*30)
            debug_print("Unknown button clicked. Start processing...")
            flash("未知的操作按钮！")
            # 若出现此情况,直接返回页面
            return render_template('AVLHandle.html',
                                   user=user,
                                   pwd=pwd,
                                   PCBA_Part_Number_List=PCBA_Part_Number_List,
                                   Version=__Version__,
                                   msg_avlHandle=msg_avlHandle,
                                   btn_enabled=btn_enabled)
        # Flask常规返回方法,因使用AJAX,此处注释掉
        # return render_template('AVLHandle.html',
        #                        user=user,
        #                        pwd=pwd,
        #                        PCBA_Part_Number_List=PCBA_Part_Number_List,
        #                        Version=__Version__,
        #                        msg_avlHandle=msg_avlHandle,
        #                        btn_enabled=btn_enabled)
        # enable buttons after processing
        btn_enabled = True
        # 对按键响应操作完成,返回JSON以便AJAX处理
        # JavaScript方式刷新页面
        return jsonify({'status': 'completed',
                         'msg': msg_avlHandle,
                         'btn_enabled': btn_enabled})
    return render_template('AVLHandle.html',
                           Version=__Version__,
                           msg_avlHandle=msg_avlHandle,
                           btn_enabled=btn_enabled)




# =========== 以下部分为Cyrus 生成的AVL BOM相关代码 ==============
#函数,功能为读取Windhill的BOM表并去除重复。输入,Excel Sheet, WinChill返回的JSON,Level是指BOM结构上的层级,1为首层
def showBOM(sheet,subpart,level):
    if level > 1:
        #判断PartNumber是否已经存在于当前AVL中
        if not subpart["PartNumber"] in AVLPart_ListView:
            #判断是否software Part或者Dcoment Part,只有不是时才往下走
            if subpart["Part"]["@odata.type"] != "#PTC.ProdMgmt.ABBDOCPART" and subpart["Part"]["@odata.type"] != "#PTC.ProdMgmt.ABBSOFTWAREPART":
                #写入到partlist里
                AVLPart_ListView.add(subpart["PartNumber"])
                #写入到Excel里
                global Excel_Row #声明Excel_Row为全局变量
                rownum1="B"+str(Excel_Row)
                rownum2="C"+str(Excel_Row)
                sheet[rownum1]=subpart["PartNumber"]    #B列写入Windchill中的PartNumber
                sheet[rownum2]=subpart["PartName"]      #C列写入Windchill中的Description
                
                #print(subpart["PartNumber"])
                #######读取CONNECT数据库并返回内容
                #根据PartNumber读取CONNECT里的内容
                print(Excel_Row-6)
                Excel_Row+=1
                ConnectFetch=db.fetchMax(subpart["PartNumber"])
                
                #声明全球变量
                global Component_ListView
                if isinstance(ConnectFetch,list):       #只有在数据库中存在该值,如果的返回的值是一个list
                    if len(ConnectFetch)>0:             #返回和数据中,超过一行,即有有效数据
                        ConnectData=ConnectFetch[0]
                        #遍历返回的CONNECT数据库

                        for index in range(1,16):
                            rownum3=chr(index+67)+str(Excel_Row-1) #第一个字母为D,从D列开始往后写
                            if ConnectData[index]!="":
                                sheet[rownum3]=ConnectData[index+2] #按照db_mgt里的数值,将结果加上与Manufactory对上的列
                        coid=Excel_Row-7
                        Component_ListView.add(str(coid)+","+subpart["PartNumber"]+","+subpart["PartName"]+","+ConnectData[1]+","+ConnectData[2])
                else:
                    Component_ListView.add(str(coid)+","+subpart["PartNumber"]+","+subpart["PartName"]+",,")                        

    #print("Components" in subpart)
    if "Components" in subpart:
        if len(subpart["Components"])>0:
            for subpart2 in subpart["Components"]:
                showBOM(sheet,subpart2,level+1)
                
                
#函数,用于检验返回的值是否Json语句,以判断是否正确地访问windchill
def is_json(myjson):  
    try:  
        json_object = json.loads(myjson)  
    except ValueError as e:  
        return False  
    return True  


#avlindex页面,生成AVL的入口页面
@app.route("/avlindex", methods=['GET','POST'])
def AVLIndex():
    return render_template('AVLIndex.html')

#avl export页面,生成AVL后的返回页面
@app.route('/exportavl',methods=['GET','POST'])  
def exportavl():
    #由于运行在服务器,每次访问时,均需要先重置Global变量以达到预期效果
    global AVLPart_ListView
    AVLPart_ListView.clear()
    global Excel_Row
    Excel_Row=7
    global Component_ListView
    Component_ListView.clear()
 
    username = request.form.get('user')
    password = request.form.get('password')
    # 将用户名和密码组合成一个字符串,并用冒号分隔  
    credentials = f"{username}:{password}"  
    # 对这个字符串进行base64编码  
    encoded_credentials = base64.b64encode(credentials.encode('utf-8'))  
    
    partnumber = request.form.get('partnumber')
    #print(partnumber)
    ########第一步,打开Excel文件并用于数据中转
    # 加载现有的 Excel 文件  
    workbook = openpyxl.load_workbook('2TFP900033A1076.xlsx')  
    #指定sheet为AVL
    sheet1 = workbook["BOM Related"]   


    ########第二步,获取WindChill Token
    url = WC_Path + 'PTC/GetCSRFToken()'  # 目标 URL  
    headers = {  
        'Authorization': 'Basic ' + encoded_credentials.decode('utf-8'),  
        'Accept': 'application/json'  
    }  
    response = requests.get(url, headers=headers)  # 发送带请求头的 GET 请求  
    #如果返回值不为JSON,重新填写
    if not is_json(response.text):
        return render_template('AVLIndex.html', ErrorMessage="访问WindChill失败,请检查用户名、密码及网络连接")
    json_data = json.loads(response.text)  
    nonce_value = json_data.get('NonceValue')  
    headers['CSRF_NONCE'] = nonce_value

    ########第三步,打开ACCESS数据库并读取AVL对应的BOM表
    #打开数据库
    bIsDBOpen=db.openAcc()
    #打开数据表,并查找对应的
    sql_result=db.readBOM(partnumber)
    
    #定义一个空的集合用于显示BOM清单
    BOM_ListView=set()
    
    #打开CONNECT数据库
    db.openMaxDB()
    
    #遍历所有sql_result
    for index in range(len(sql_result)):
        
        
        # 修改单元格内容
        rownum = 'B' + str(index+3)
        print(index)
        BomNum = re.sub(r'[\',()]', '', str(sql_result[index])) 
        sheet1[rownum] = BomNum
        
        ########第四步,从WindChill里导入BOM表的状态
        url = WC_Path + f"ProdMgmt/Parts?$filter=Number eq '{BomNum}'"  # 目标 URL  
        response = requests.get(url, headers=headers)  # 发送带请求头的 GET 请求  
        json_data = json.loads(response.text)  
        partID=""
        # 遍历value数组中的每个元素  
        for partvalue in json_data['value']:  
            # 检查'View'字段是否为'design'（不区分大小写）  
            if partvalue['View'].lower() == 'design':  
                # 创建一个字典来存储part的详细信息  
                BOM_ListView.add(BomNum+","+ partvalue['State']['Value']+","+ partvalue['Version']+","+  partvalue['Name'])
                partID = partvalue["ID"]
                rownum = 'C' + str(index+3)
                sheet1[rownum] = partvalue['Name']
                # 由于已经找到了匹配的'design'视图,因此跳出循环  
                break
        
        #第五步,根据第四步返回的ID,查找对应的BOM结构
        url= WC_Path + "ProdMgmt/Parts('" + partID + "')/PTC.ProdMgmt.GetBOM?$expand=Components($expand=Part($select=Name,Number);$levels=max)"
        response = requests.post(url, headers=headers)  # 发送带请求头的 GET 请求 
        json_data = json.loads(response.text)          
        showBOM(workbook["AVL"],json_data,1)
    
    
    
    # 保存修改后的工作簿  
    workbook.save(filename="out/modified_example.xlsx")      
    PartCount=Excel_Row-7
    
    return render_template('AVLoutput.html', sql_result=BOM_ListView,AVL=partnumber,PartCount=PartCount,componentlist=Component_ListView)

#超链接,用于下载相应的Excel文件
@app.route('/downloadexcel/<AVL>')  
def downloadFile(AVL):  
    # 返回修改后的Excel文件供下载  
    modified_file = open("out/modified_example.xlsx", "rb")  
    
    # 保存并准备下载  
    response = send_file(modified_file, download_name=AVL+'.xlsx', as_attachment=True)  
    
    return response

if __name__ == '__main__':  
    app.run(host="0.0.0.0", debug = True)





