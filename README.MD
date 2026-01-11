# 开发环境： 
    Python + Flask + PyPika

# 运行环境：
##    Python 3.9+  
##   add-on:   
      Flask 2.0.2  
      PyPika 0.48.9  
   实测未能运行在IIS环境下，需要单独通过Python程序运行。

# 程序功能： 
   通过ODBC驱动读取CONNECT DB数据信息，支持Cadence CIS数据库和Access数据。  
   可以自行选择使用哪个数据库进行读取。  
   需先选择器件类型，然后设置搜索条件进行查找； 也可以搜索全部器件类型。  
   这个应用比CONNECT系统和AD软件中搜索速度更快，并且可以支持多种组合条件搜索，搜索结果可以导出。
	

# Update Log:(only for the Production version)
   20240129: 增加自动AVL生成功能，用于电子组未来所有AVL均由此生成再上传到WindChill。  
     → by Cyrus。2025之后因PLM改为双重验证登录，此功能已无法使用，需更新。  
   20240407: 添加检索Manufacturer Part Num的功能并实现。  
   20240407: 添加导出检索结果到Excel文件的功能。  
   20260104: 将Dev_AVL_Cyrus分支合并回master，以后的功能开发基于master进行。  
   20260108: 重构代码结构，db_mgt.py使用PyPika进行SQL语句生成，提升代码可维护性和扩展性；并实现搜索条件AND运算。  
   20260111: 对应2.5.0版本。  
     a. 新增过滤条件“SAP_Description”、“techdescription”、“editor”，网页端增加输入框。  
     b. 新增多SAP_Num搜索功能，可以一次搜索多个物料并保存Excel文件，用于输出文件作AVL对比很方便。  

# git branch:
1. master:   
   major branch, 
2. dev:   
   developing branch, new feature develop here. This branch will be mergedback to "master".
3. Dev_AVL_Cyrus:   
   Cyrus add AVL Export f unction web page in the server. first version of this branch is a copy from the server. After then, "dev" branch update will be merged to this branch before release to the server. This Branch had been merged back to "master" on 20260104 and will be no longer used.
4. dev_IIS:  
   IIS environment refactor branch.
   To let the code can run in IIS + WSGI
   --> access permission problem, pending.
5. Refactor_PyPika  
   use PyPika lib to refactor the SQL generate method.
6. Production  
   branch for production.  
   The server can sync this branch directly in production environment.  

# ToDo List:
1. 显示窗口有点难看，需要使用CCS进行润色 (Done)  
    index.html页面添加ccs代码格式化表格显示，表头可固定  
    未找到固定第一列的方法
2. Part Type中虽然有显示"---ALL----"的选项，但是未支持一次搜索所有Table. (Done)  
   --> 已实现所有Table的基本搜索功能，但只显示有限的select_fields = 'PartNumber,value_1,SAP_Number,SAP_Description'  
   因为有些Table不带有其它的属性。  
   --> 下一步：修改当进行全Table搜索时，Table list仅包括有关注参数的那些Table，我们需要输出的列除上述外还需要，Symbol, Footprint信息， Ordering info信息等共有的信息。  
   --> Done
3. 搜索条件仅支持自上而下的一种，不支持多种组合 -> 暂未有计划
4. Search All时显示drawing字段信息 (Done)  
    -> access db实现“detaildrawing”  
    -> CIS DB实现“detaildrawing”  
5. 增加直接访问DESTO CIS DB的选项，要求先在服务器端配置相应的DSN (Done)
   在ODBC中添加新的设置,信息如下：  
      Driver:SAP MaxDB  
      Name:CIS_DESTO  
      Description:CONNECT DB for CDS DESTO online  
      Server:xc-s-zl08506.xc.abb.com  
      Database:limbas  
   程序中添加“04-CONNECT DESTO(ODBC)”数据库选择，并实现。
6. 将"01-CONNECT Online(ODBC)"更换为"01-CONNECT Local(ODBC)" (Done)
   并切换ODBC为"CIS_Local",可以避免因服务器因使用Cadence时切换"CONNECT Partslib V2"导致程序出现问题。
   在ODBC中添加新的设置,信息如下  
      Driver:SAP MaxDB  
      Name:CIS_Local  
      Description:CONNECT DB for CDS Local  
      Server:cn-w-7295178.cn.abb.com  
      Database:limbas  
7. 对检索的结果进行排序 (Done)
   sql_fetch = sql_fetch + " ORDER BY PartNumber ASC" 
8. IIS WSGI部署 (Pending) 
   依然存在的问题  
   a) UNC连接到P盘源目录，但是它的权限没有对IIS的用户配置开放，依然无法访问  
   b) log需要移植最新的easy_log过来  
   --> 使用IIS运行存在非常多的权限问题需要解决，暂时不继续搞了
   

