'''
说明：
使用Flask访问CONNECT DB并显示
即制作网页版的CONNECT DB显示程序

阶段实现说明：
1. 已经实现原来Tkinter程序的功能：\01_Prg\10_Py\PythonCode\L_DB
2. 设置最大显示数量是因为如果检索结果数量过多，渲染表格太大显示有问题。数据检索的条目不作限制，仅是限制了显示的数量。

todo:
1. 整个页面的样式太难看了，需要使用合适的ccs。
2. Production系统如何Debug？

?:
1. 网页端选择了DB Selection下拉框之后，Part Type的列表没有办法更新，这是需要js来实现的吧？
==>需要先有一个选择数据库的页面，而后跳转到新的检索页面，并且已经确认使用哪个数据。
==>已经实现

'''

from flask import Flask, request, redirect
from flask import render_template
from flask.helpers import flash, url_for
import db_mgt
import logging
import openpyxl
import tempfile
import os

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

# 对于IIS生产系统，这段要放在这里，不能放在main里
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

'''
初始始页面： 选择数据库
点击确认之后，会跳转到检索页面，并且会将选择的数据库序号传递过去。
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
        return render_template('DBSelect.html', DB_List=DB_List,)

'''
数据库检索和显示页面：初始始显示检索表单，在此之前会尝试打开数据库
submit之后显示检索内容
'''
@app.route("/search/<DBType>", methods=['GET','POST'])
def index(DBType):
    # 根据DBType来设置Part Type 列表的内容，DBType为str，对应db_mgt.DBList的index值，从0开始
    if DBType == '0' or DBType == '3': 
        #如果将值直接在render_template里赋值，数据第一次会传递不过去，不知原因。
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
        # 判断按键
        print("press the button:" + request.form['btn'])
        if request.form['btn'] == 'Search':
            # 获取检索条件
            global sql_result
            global columnNameList
            global MaxLine
            global sql_result_len
            PartNo_Searchby = request.form.get("PartNo")
            SAPNo_Searchby = request.form.get("SAPNo")
            PartValue_Searchby = request.form.get("PartValue")
            MfcPartNum_Searchby = request.form.get("MfcPartNum")
            MaxLine = int(request.form.get("MaxLine"))
            tableName  = request.form.get("tableName")
            dbindex = int(DBType)
            sql_result, columnNameList = db.fetch(tableName, dbindex, PartNo_Searchby, SAPNo_Searchby, PartValue_Searchby, MfcPartNum_Searchby)
            sql_result_len = len(sql_result)
            return render_template('index.html', Part_Type_List=Part_Type_List, MaxLine=MaxLine, sql_result=sql_result, columnNameList=columnNameList, sql_result_len=sql_result_len)
            # return db_mgt.DBList[0]
        elif request.form['btn'] == 'SaveExcel':
            # 保存Excel
            print("SaveExcel")
            # print(columnNameList)
            if 'columnNameList' in globals():
                temp_dir = tempfile.gettempdir()
                file_path = os.path.join(temp_dir,"SQL_Result.xlsx")
                if file_path:
                    wb = openpyxl.Workbook()
                    ws = wb.active
                    ws.append(columnNameList)
                    for row in sql_result:
                        ws.append(row)
                    wb.save(file_path)
                    flash("Excel保存成功！{}".format(file_path))
                # 打开Excel
                os.system('start excel.exe {}'.format(file_path))
                return render_template('index.html', Part_Type_List=Part_Type_List, MaxLine=MaxLine, sql_result=sql_result, columnNameList=columnNameList, sql_result_len=sql_result_len)
            else:
                flash("没有数据，无法保存Excel！")
                return render_template('index.html', Part_Type_List=Part_Type_List)
        else:
            return render_template('index.html', Part_Type_List=Part_Type_List)
    else:
        return render_template('index.html', Part_Type_List=Part_Type_List)

if __name__ == '__main__':

    
    app.run(host="0.0.0.0", debug = True)
