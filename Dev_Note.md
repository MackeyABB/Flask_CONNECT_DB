# 说明：
此文件用于保存开发过程中发现的bug、任务、idea快速记录与跟进

# Bug:
1. app.py - 4.2.0 done
以下程序中"DBType == '3'"是AccessDB, 不应该跟1放在一起的，这导致选择这个数据库时，
在Table选择下拉框中导致错误的TableName形式，需要在正式版中进行修改。

def index(DBType):
    db = get_db()
    # 根据DBType来设置Part Type 列表的内容,DBType为str,对应db_mgt.DBList的index值,从0开始
    if DBType == '0' or DBType == '3': 


# Task todo:
1. PostgreDB 使用需要对SQL生成重新编写 - 4.2.0 done
TableName，FieldName使用""包含
FieldName需要严格遵守大小写
2. 更新"AVL handle Page"以支持PostgreSQL DB - 4.3.0 done
当前仅支持"CONNECT DB""Access DB两个选项"，需要增加添加新数据库的支持。


# idea:
