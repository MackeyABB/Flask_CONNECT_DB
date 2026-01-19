'''
Introduction:
    此程序用于通过Basic Auth方式，绕过Windchill的MFA验证，访问Windchill的OData API，获取指定物料号的BOM结构信息。
    主要功能包括：
    1. 生成Basic Auth Token
    2. 获取Windchill的CSRF Token并构造认证请求头
    3. 根据物料号查询BOM ID
    4. 获取并解析BOM结构，支持仅获取中国物料号或全部物料号
Usage:
    1. 设置环境变量 PLM_User 和 PLM_Pw，分别为Windchill的用户名和密码。
    2. 运行脚本，脚本会输出指定物料号的BOM清单。
    3. 此文件可以被其它脚本导入，调用 get_bom 函数获取BOM信息。
Notes:

Revision log:
1.0.0 - 20260114: 创建文件，实现基本功能。
1.1.0 - 20260114: 增加返回数据的类型，包括SAP_Number_List和对应的string格式，方便后续处理。
1.2.0 - 20260115: 修改返回数据结构，set改为list类型，更符合实际用途。
1.3.0 - 20260115: 增加函数get_BOM(user, pwd, BOM_number, IsChinaPNOnly=bCHINA_PN_ONLY)可以一步获取BOM信息，简化调用。
1.4.0 - 20260119: 优化PLM登录失败的处理逻辑，避免后续函数调用出错。get_BOM()函数增加返回PLM登录是否成功的标志PLM_Login_OK。
'''

# 版本号
# xx.yy.zz
# xx: 大版本，架构性变化
# yy: 功能性新增
# zz: Bug修复
__revision__ = '1.4.0'

import requests
import base64
import json
import os

# global variables
# 以下链接是bypass MFA的API
PLM_Bypass_MFA_Url = "https://lp-global-plm.abb.com/Windchill/protocolAuth/servlet/odata/"


#是否为仅中国物料
bCHINA_PN_ONLY = True
bALL_PN = False

# functions
def is_json(myjson):  
    '''检查字符串是否为有效的JSON格式
    Args:
        myjson (str): 要检查的字符串
    Returns:
        bool: 如果是有效的JSON格式则返回True，否则返回False
    '''
    try:  
        json_object = json.loads(myjson)  
    except ValueError as e:  
        return False  
    return True  

def get_basic_token(user, pwd):
    '''生成Basic Auth Token
    Args:
        user (str): 用户名
        pwd (str): 密码
    Returns:
        str: Basic Auth Token
    '''
    raw = f"{user}:{pwd}".encode("utf-8")
    return base64.b64encode(raw).decode("ascii")

def generate_auth_header(token):
    '''获取WindChill Token
    Args:
        token (str): Basic Auth Token
    Returns:
        dict: 包含认证信息的请求头
    参考json数据实例参照文件：../Debug/PLM_Get_Auth_header.json
    '''
    url = PLM_Bypass_MFA_Url + 'PTC/GetCSRFToken()'  # 目标 URL  
    headers = {
            "Authorization": f"Basic {token}",
            'Accept': 'application/json'
        }
    response = requests.get(url, headers=headers)  # 发送带请求头的 GET 请求  
    #如果返回值不为JSON，重新填写
    if not is_json(response.text):
        print("访问WindChill失败，请检查用户名、密码及网络连接")
        return None
    json_data = json.loads(response.text)  
    nonce_value = json_data.get('NonceValue')  
    headers['CSRF_NONCE'] = nonce_value
    return headers

def get_parse_Uses_BOM(json_data, BOM_number, IsChinaPNOnly=bCHINA_PN_ONLY):
    '''解析Uses BOM结构
    Args:
        json_data (dict): 包含BOM结构的JSON数据
        BOM_number (str): 物料号
        IsChinaPNOnly (bool): 是否仅获取中国物料号，默认为True
    Returns:
        BOM_Info_list (list): 包含BOM清单的列表, 每个元素格式为 "PartNumber,PartName,Quantity,DesignatorRange"
        BOM_SAP_Number_List (list): 包含BOM的SAP_Number清单的列表
        BOM_SAP_Number_List_Str (str): 包含BOM的SAP_Number清单
    note:
        1. 仅解析Uses BOM结构
        2. 仅在IsChinaPNOnly为True时，才会过滤出中国物料号
    '''
    # structure:
    # "Uses": [  # 一级组件列表
    #    {"ReferenceDesignatorRange": "R20,R36",
    #     "Quantity": 2,
    #     "Uses": [  # 二级组件列表
    #         {"Number": "2TFUxxxx"
    #         "Name": "xxxx"
    #         ...},
    #         ...   
    # joson数据实例参照文件：../Debug/PLM_2TFE001045D1801_BOM_Structure.json
    # 定义一个空的list用于保存BOM清单
    BOM_Info_list = []
    # 定义一个list和str用于存储BOM的SAP_Number清单
    BOM_SAP_Number_List = []
    BOM_SAP_Number_List_Str = ""
    for component in json_data.get("Uses", []):
        PartNumber = component.get("Uses", {}).get("Number", "")
        PartName = component.get("Uses", {}).get("Name", "")
        Quantity = component.get("Quantity", 0)
        DesignatorRange = component.get("ReferenceDesignatorRange", "")
        # need to remove the PCBA part
        if BOM_number != PartNumber: 
            if IsChinaPNOnly:
                # add only China PN starting with 2TFU to the BOM_ListView
                if PartNumber.startswith("2TFU"):
                    BOM_Info_list.append(f"{PartNumber},{PartName},{Quantity},{DesignatorRange}")
                    BOM_SAP_Number_List.append(PartNumber)
                    BOM_SAP_Number_List_Str += PartNumber + ","
            else:
                # all BOM items to the BOM_ListView, should include some other documents parts.
                BOM_Info_list.append(f"{PartNumber},{PartName},{Quantity},{DesignatorRange}")
                BOM_SAP_Number_List.append(PartNumber)
                BOM_SAP_Number_List_Str += PartNumber + ","
        else:
            print(f"跳过自身物料号: {PartNumber}")
    return BOM_Info_list, BOM_SAP_Number_List, BOM_SAP_Number_List_Str

def get_Uses_Bom(headers, BOM_number, IsChinaPNOnly=bCHINA_PN_ONLY):
    '''获取指定物料号的BOM结构
    Args:
        headers (dict): 包含认证信息的请求头
        BOM_number (str): 物料号
        IsChinaPNOnly (bool): 是否仅获取中国物料号，默认为True
    Returns:
        parse_BOM (set): 包含BOM清单的集合, 每个元素格式为 "PartNumber,PartName,Quantity,DesignatorRange"
        BOM_SAP_Number_List (list): 包含BOM的SAP_Number清单的列表
        BOM_SAP_Number_List_Str (str): 包含BOM的SAP_Number清单
        PCBA_Part_info_list (list): 包含PCBA Part的详细信息
    note:
        1. 获取Uses BOM结构
        2. 仅在IsChinaPNOnly为True时，才会过滤出中国物料号
        3. 需要调用另外的函数 get_parse_Uses_BOM 来解析BOM结构
    参考json数据实例参照文件：../Debug/PLM_PartNumber_value.json
    '''
    if headers is None:
        return None
    # step1: 获取BOM ID
    url = PLM_Bypass_MFA_Url + f"ProdMgmt/Parts?$filter=Number eq '{BOM_number}'"
    response = requests.get(url, headers=headers)  # 发送带请求头的 GET 请求  
    json_data = json.loads(response.text)  
    partID=""
    # 遍历value数组中的每个元素  
    # joson数据实例参照文件：../Debug/PLM_PartNumber_value.json
    # 定义一个空的list用于保存PCBA Part的详细信息
    PCBA_Part_info_list=[]  
    for partvalue in json_data['value']:  
        # 检查'View'字段是否为'design'（不区分大小写）, 与manufacture区分
        if partvalue['View'].lower() == 'design':  
            # 存储part的详细信息  
            PCBA_Part_info_list.append(BOM_number)
            PCBA_Part_info_list.append(partvalue['State']['Value'])
            PCBA_Part_info_list.append(partvalue['Version'])
            PCBA_Part_info_list.append(partvalue['Name'])
            partID = partvalue["ID"]
            # 由于已经找到了匹配的'design'视图，因此跳出循环  
            break
    
    # Step2: 根据partID，查找对应的BOM结构
    # 可以先通过以下获取精确属性名，然后构造查询
    # https://lp-global-plm.abb.com/Windchill/protocolAuth/servlet/odata/ProdMgmt/$metadata
    # 返回数据显示 Windchill 系统的 OData 结构，核心是 Part 实体没有 BOMView，但提供了专用的 BOM 相关实体和动作来获取结构。
    ############  获取BOM结构的几种方案 ############
    # 方案 1：通过 Uses 导航属性获取一级 BOM（最直接）
    # https://lp-global-plm.abb.com/Windchill/protocolAuth/servlet/odata/ProdMgmt/Parts('OR:wt.part.WTPart:43943079683')?$expand=Uses($expand=Uses)    
    url= PLM_Bypass_MFA_Url + "ProdMgmt/Parts('" + partID + "')?$expand=Uses($expand=Uses)"  # 获取一级BOM结构
    # 方案 2：方案 2：通过专用动作 GetPartStructure 获取多层级 BOM（推荐）
    # 元数据中定义了 GetPartStructure 动作，是 Windchill 官方推荐的获取多层级 BOM 结构的方式，支持完整的层级展开：
    # https://lp-global-plm.abb.com/Windchill/protocolAuth/servlet/odata/ProdMgmt/Parts('OR:wt.part.WTPart:43943079683')/PTC.ProdMgmt.GetPartStructure
    # Todo: 也出错，需要找原因
    # url= PLM_Bypass_MFA_Url + "ProdMgmt/Parts('" + partID + "')/PTC.ProdMgmt.GetPartStructure?$expand=Components($expand=Part($select=Name,Number);$levels=max)"
    # 方案 3： 这是参考Cyrus程序的方法，出错
    # url= PLM_Bypass_MFA_Url + "ProdMgmt/Parts('" + partID + "')/PTC.ProdMgmt.GetBOM?$expand=Components($expand=Part($select=Name,Number);$levels=max)"    
    response = requests.get(url, headers=headers)  # 发送带请求头的 GET 请求 
    json_data = json.loads(response.text)    
    print("BOM结构获取成功")
    # Step3: 解析BOM结构
    BOM_Info_list, BOM_SAP_Number_List, BOM_SAP_Number_List_Str = get_parse_Uses_BOM(json_data, BOM_number,  IsChinaPNOnly)
    return BOM_Info_list, BOM_SAP_Number_List, BOM_SAP_Number_List_Str, PCBA_Part_info_list

def get_BOM(user, pwd, BOM_number, IsChinaPNOnly=bCHINA_PN_ONLY):
    '''一步获取BOM信息，简化调用。
    Args:
        user (str): 用户名
        pwd (str): 密码
        BOM_number (str): 物料号
        IsChinaPNOnly (bool): 是否仅获取中国物料号，默认为True
    Returns:
        BOM_Info_list (list): 包含BOM清单的列表, 每个元素格式为 "PartNumber,PartName,Quantity,DesignatorRange"
        BOM_SAP_Number_List (list): 包含BOM的SAP_Number清单的列表
        BOM_SAP_Number_List_Str (str): 包含BOM的SAP_Number清单
        PCBA_Part_info_list (list): 包含PCBA Part的详细信息
        PLM_Login_OK (bool): PLM登录是否成功
    '''
    token = get_basic_token(user, pwd)
    headers = generate_auth_header(token)
    if headers is None:
        PLM_Login_OK = False
        BOM_Info_list, BOM_SAP_Number_List, BOM_SAP_Number_List_Str, PCBA_Part_info_list = [], [], "", []
    else:
        PLM_Login_OK = True
        BOM_Info_list, BOM_SAP_Number_List, BOM_SAP_Number_List_Str, PCBA_Part_info_list = get_Uses_Bom(headers, BOM_number, IsChinaPNOnly)
    return BOM_Info_list, BOM_SAP_Number_List, BOM_SAP_Number_List_Str, PCBA_Part_info_list, PLM_Login_OK

if __name__ == "__main__":
    if True:
        # 从环境变量获取用户名和密码
        # 若出现问题，可以使用以下方式打开vscode
            # setx PLM_User "你的PLM用户名"  # 若已配置，可跳过这步，直接启动VS Code
            # setx PLM_Pw "你的PLM密码"
            # code  # 启动VS Code（需确保code命令已加入环境变量，若没生效，用VS Code的完整路径）
        user=os.getenv("PLM_User")
        pwd=os.getenv("PLM_Pw")
        # user="CNMALAO"
        # pwd="tbd"
        # 校验变量是否存在
        if not user:
            print("警告：环境变量 PLM_User 未设置！")
        if not pwd:
            print("警告：环境变量 PLM_Pw 未设置！")
    else:
        user=input("input user(short name, e.g.:CNMALAO):")
        pwd=input("input pwd:")
    token = get_basic_token(user, pwd)
    headers = generate_auth_header(token)
    BOM_number = "2TFE001045D1801"
    BOM_Info_list, BOM_SAP_Number_List, BOM_SAP_Number_List_Str, PCBA_Part_info_list = get_Uses_Bom(headers, BOM_number, bALL_PN)
    print("BOM清单如下：")
    for item in BOM_Info_list:
        print(item)
    for item in BOM_SAP_Number_List:
        print(item)
    print(BOM_SAP_Number_List_Str)
    print("PCBA Part信息如下：")
    for item in PCBA_Part_info_list:
        print(item)
    
    # 也可以直接调用 get_BOM 函数, 直接获取BOM信息
    BOM_Info_list, BOM_SAP_Number_List, BOM_SAP_Number_List_Str, PCBA_Part_info_list = get_BOM(user, pwd, BOM_number, bCHINA_PN_ONLY)
    print("="*40)
    print("通过 get_BOM 函数获取的BOM清单如下：")
    for item in BOM_Info_list:
        print(item)
    for item in BOM_SAP_Number_List:
        print(item)
    print(BOM_SAP_Number_List_Str)
    print("通过 get_BOM 函数获取的PCBA Part信息如下：")
    for item in PCBA_Part_info_list:
        print(item)
    