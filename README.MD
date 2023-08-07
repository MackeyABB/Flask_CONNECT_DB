开发环境： 
    Pythone + Flask

程序功能： 
    通过ODBC驱动读取数据CONNECT DB，支持Cadence CIS数据库和Access数据。
    可以自行选择使用哪个数据库进行读取。
    需先选择器件类型，然后设置搜索条件进行查找。


ToDo List:
1. 显示窗口有点难看，需要使用CCS进行润色 -> 暂未有计划
2. Part Type中虽然有显示"---ALL----"的选项，但是未支持一次搜索所有Table. (Done)
   --> 已实现所有Table的基本搜索功能，但只显示有限的select_fields = 'PartNumber,value_1,SAP_Number,SAP_Description'
   因为有些Table不带有其它的属性。
   --> 下一步：修改当进行全Table搜索时，Table list仅包括有关注参数的那些Table，我们需要输出的列除上述外还需要，Symbol, Footprint信息， Ordering info信息等共有的信息。
   --> Done
3. 搜索条件仅支持自上而下的一种，不支持多种组合 -> 暂未有计划



