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

# create DB instance
WC_Path = "https://lp-global-plm.abb.com/Windchill/protocolAuth/servlet/odata/"
#定义一个空的集合用于记录AVL里的元器件清单
AVLPart_ListView=set()
#定义一个空的集合用于在网页端显示内容以供使用者检查
Component_ListView=set()
#定义Excel模板中，有效数据的首行
Excel_Row=7


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
        # 获取检索条件
        PartNo_Searchby = request.form.get("PartNo")
        SAPNo_Searchby = request.form.get("SAPNo")
        PartValue_Searchby = request.form.get("PartValue")
        MaxLine = int(request.form.get("MaxLine"))
        tableName  = request.form.get("tableName")
        dbindex = int(DBType)
        sql_result, columnNameList = db.fetch(tableName, dbindex, PartNo_Searchby, SAPNo_Searchby, PartValue_Searchby)
        sql_result_len = len(sql_result)
        return render_template('index.html', Part_Type_List=Part_Type_List, MaxLine=MaxLine, sql_result=sql_result, columnNameList=columnNameList, sql_result_len=sql_result_len)
        # return db_mgt.DBList[0]
    else:
        return render_template('index.html', Part_Type_List=Part_Type_List)


#函数，功能为读取Windhill的BOM表并去除重复。输入，Excel Sheet, WinChill返回的JSON，Level是指BOM结构上的层级，1为首层
def showBOM(sheet,subpart,level):
    if level > 1:
        #判断PartNumber是否已经存在于当前AVL中
        if not subpart["PartNumber"] in AVLPart_ListView:
            #判断是否software Part或者Dcoment Part，只有不是时才往下走
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
                if isinstance(ConnectFetch,list):       #只有在数据库中存在该值，如果的返回的值是一个list
                    if len(ConnectFetch)>0:             #返回和数据中，超过一行，即有有效数据
                        ConnectData=ConnectFetch[0]
                        #遍历返回的CONNECT数据库

                        for index in range(1,16):
                            rownum3=chr(index+67)+str(Excel_Row-1) #第一个字母为D,从D列开始往后写
                            if ConnectData[index]!="":
                                sheet[rownum3]=ConnectData[index+2] #按照db_mgt里的数值，将结果加上与Manufactory对上的列
                        coid=Excel_Row-7
                        Component_ListView.add(str(coid)+","+subpart["PartNumber"]+","+subpart["PartName"]+","+ConnectData[1]+","+ConnectData[2])
                else:
                    Component_ListView.add(str(coid)+","+subpart["PartNumber"]+","+subpart["PartName"]+",,")                        

    #print("Components" in subpart)
    if "Components" in subpart:
        if len(subpart["Components"])>0:
            for subpart2 in subpart["Components"]:
                showBOM(sheet,subpart2,level+1)
                
                
#函数，用于检验返回的值是否Json语句，以判断是否正确地访问windchill
def is_json(myjson):  
    try:  
        json_object = json.loads(myjson)  
    except ValueError as e:  
        return False  
    return True  

#avlindex页面，生成AVL的入口页面
@app.route("/avlindex", methods=['GET','POST'])
def AVLIndex():
    return render_template('AVLIndex.html')

#avl export页面，生成AVL后的返回页面
@app.route('/exportavl',methods=['GET','POST'])  
def exportavl():
    #由于运行在服务器，每次访问时，均需要先重置Global变量以达到预期效果
    global AVLPart_ListView
    AVLPart_ListView.clear()
    global Excel_Row
    Excel_Row=7
    global Component_ListView
    Component_ListView.clear()
 
    username = request.form.get('user')
    password = request.form.get('password')
    # 将用户名和密码组合成一个字符串，并用冒号分隔  
    credentials = f"{username}:{password}"  
    # 对这个字符串进行base64编码  
    encoded_credentials = base64.b64encode(credentials.encode('utf-8'))  
    
    partnumber = request.form.get('partnumber')
    #print(partnumber)
    ########第一步，打开Excel文件并用于数据中转
    # 加载现有的 Excel 文件  
    workbook = openpyxl.load_workbook('2TFP900033A1076.xlsx')  
    #指定sheet为AVL
    sheet1 = workbook["BOM Related"]   


    ########第二步，获取WindChill Token
    url = WC_Path + 'PTC/GetCSRFToken()'  # 目标 URL  
    headers = {  
        'Authorization': 'Basic ' + encoded_credentials.decode('utf-8'),  
        'Accept': 'application/json'  
    }  
    response = requests.get(url, headers=headers)  # 发送带请求头的 GET 请求  
    #如果返回值不为JSON，重新填写
    if not is_json(response.text):
        return render_template('AVLIndex.html', ErrorMessage="访问WindChill失败，请检查用户名、密码及网络连接")
    json_data = json.loads(response.text)  
    nonce_value = json_data.get('NonceValue')  
    headers['CSRF_NONCE'] = nonce_value

    ########第三步，打开ACCESS数据库并读取AVL对应的BOM表
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
        
        ########第四步，从WindChill里导入BOM表的状态
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
                # 由于已经找到了匹配的'design'视图，因此跳出循环  
                break
        
        #第五步，根据第四步返回的ID，查找对应的BOM结构
        url= WC_Path + "ProdMgmt/Parts('" + partID + "')/PTC.ProdMgmt.GetBOM?$expand=Components($expand=Part($select=Name,Number);$levels=max)"
        response = requests.post(url, headers=headers)  # 发送带请求头的 GET 请求 
        json_data = json.loads(response.text)          
        showBOM(workbook["AVL"],json_data,1)
    
    
    
    # 保存修改后的工作簿  
    workbook.save(filename="out/modified_example.xlsx")      
    PartCount=Excel_Row-7
    
    return render_template('AVLoutput.html', sql_result=BOM_ListView,AVL=partnumber,PartCount=PartCount,componentlist=Component_ListView)

#超链接，用于下载相应的Excel文件
@app.route('/downloadexcel/<AVL>')  
def downloadexcel(AVL):  
    # 返回修改后的Excel文件供下载  
    modified_file = open("out/modified_example.xlsx", "rb")  
    
    # 保存并准备下载  
    response = send_file(modified_file, download_name=AVL+'.xlsx', as_attachment=True)  
    
    return response

if __name__ == '__main__':  
    app.run(host="0.0.0.0", debug = True)
    
    



    