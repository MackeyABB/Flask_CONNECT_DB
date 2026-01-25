
# Revision Log

| 版本号 | 更新日期 | 版本说明 |
|--------|----------|----------|
| 1.0.0  | 2024-04-22 | 初始版本,实现CONNECT DB的网页端查询和Excel保存功能 |
| 1.0.1  | 2024-04-29 | 修正了Excel保存功能的bug,之前保存的Excel文件无法打开 |
| 1.1.0  | 2024-05-06 | 增加了对Access DB的支持,可以选择CONNECT DB和Access DB进行查询 |
| 1.1.1  | 2024-05-10 | 修正了Access DB查询时的bug,之前查询结果不正确 |
| 1.2.0  | 2024-05-15 | 优化了网页界面,增加了查询条件的输入框 |
| 1.2.1  | 2024-05-20 | 修正了网页界面的一些显示问题,提升用户体验 |
| 1.3.0  | 2024-05-25 | 增加了对AVL BOM导出的支持,可以从Windchill获取BOM并生成Excel文件, by Cyrus |
| 2.0.0  | 2026-01-08 | 重构代码结构,db_mgt.py使用PyPika进行SQL语句生成,提升代码可维护性和扩展性 |
| 2.1.0  | 2026-01-08 | 将搜索的条件显示在页面 |
| 2.2.0  | 2026-01-08 | 搜索条件输入内容在点击search之后不会清除 |
| 2.3.0  | 2026-01-08 | 新增过滤条件“SAP_Description”"techdescription" "editor", 网页端增加输入框。 |
| 2.4.0  | 2026-01-08 | 支持多个SAP编号的批量查询,输入多个SAP编号,按空格、逗号、分号分隔, 未完成保存Excel功能 |
| 2.5.0  | 2026-01-11 | 完成多个SAP编号的批量查询结果直接保存为Excel文件功能 |
| 2.6.0  | 2026-01-11 | 在网页端显示软件版本号 |
| 3.0.0  | 2026-01-18 | 增加了AVL处理页面,支持从Windchill获取BOM,查询数据库,生成AVL Excel文件,并通过AJAX方式下载文件<br>a) 网站打开首页增加跳转到AVL处理页面的按钮, 并添加版本号显示<br>b) 新增AVL处理页面,支持Create AVL和Download AVL功能,使用AJAX方式处理请求和下载文件<br>c) AVL处理页面中的AVL inlcude选项支持“2TFU CN only”和“All”,默认为“2TFU CN only”, 但All选项还存在问题,需要后续修正 |
| 3.1.0  | 2026-01-18 | AVL页面添加跳转回主页面的按钮 |
| 3.1.1  | 2026-01-19 | 修正了AVL处理页面中的bug: AVL_include选项为"All Parts"时,未正确输出找不到ordering information的Parts导出Excel文件的问题。 |
| 3.2.0  | 2026-01-19 | 优化了AVL处理页面的问题, 如果输入Windchill用户名和密码为空, PCBA part number为空, 则提示并不继续处理 |
| 3.3.0  | 2026-01-19 | 增加判断是否获取到ordering information, 若没有则不继续处理,并提示用户 |
| 3.3.1  | 2026-01-19 | 优化PLM登录失败的处理逻辑,避免后续函数调用出错。<br>PLM_Basic_Auth_ByPass_MFA_Get_BOM.py升级到1.4.0版本,get_BOM()函数增加返回PLM登录是否成功的标志PLM_Login_OK。 |
| 3.3.2  | 2026-01-19 | 修正CONNECT Viewer页面中的SAP Number List Search功能的bug<br>当SAP编号未找到时,会进行判断，并添加空行占位,填写SAP number |
| 3.4.0  | 2026-01-19 | "Download_AVL"按键实现下载功能 |
| 3.5.0  | 2026-01-21 | 增加AVL Comparison功能,支持上传手动整理的AVL文件进行对比,并生成对比结果Excel文件供下载。此功能暂时不支持自动生成AVL_Cmp sheet,需要用户手动整理后上传进行对比。临时版本号提升为3.5.0,等待后续完善自动生成AVL_Cmp sheet功能。 |
| 3.6.0  | 2026-01-21 | 增加Compare_Manual_AVL按键, 用于上传手动整理的AVL文件进行对比,并生成对比结果Excel文件供下载。 |
| 3.7.0  | 2026-01-21 | 实现Compare_AVL按键的AVL_Sheet_Only功能,根据上传Excel文件中的"AVL" sheet内容,查询数据库获取ordering information,并与"AVL" sheet内容进行对比,生成对比结果Excel文件供下载。<br>输出文件命名规则调整为: 原文件名_时间戳.xlsx, 方便区分不同的对比结果文件。 |
| 3.8.0  | 2026-01-21 | 实现Compare_AVL按键的BOM_Related_Sheet功能<br>根据上传Excel文件中的"BOM Related" sheet内容,获取PCBA Part Numbers列表,通过PLM获取BOM中的SAP Numbers,查询数据库获取ordering information,并与"BOM Related" sheet中的AVL内容进行对比,生成对比结果Excel文件供下载。<br>此部分代码是可以跟Create AVL部分代码进行重构复用的, 但是为了避免影响已经稳定运行的Create AVL功能, 暂时不进行重构。 |
| 3.8.1  | 2026-01-23 | 修改页面, 提示用户名不支持邮箱格式登录 |
| 3.9.0  | 2026-01-23 | 整理AVL Hanle页面, 能在较小的垂直分辨率下也能一屏显示更多内容 |
| 3.10.0 | 2026-01-24 | 修改excel_mapping, 增加Detail Drawing, AVL Status, Editor, Technical Description字段的映射<br>first_write_AVL_to_excel()函数中增加对应字段的写入,以便用户检查CONNECT中Part的状态 |
| 3.11.0  | 2026-01-24 | 版本号添加超链接可以打开Revision_Log.md文件以显示版本历史信息<br>需要安装markdown插件 |
| 3.12.0  | 2026-01-24 | 更改DBList中四个数据库的描述名称, 使其更清晰易懂。<br>修改后：<br>    第一个为直连DESTO Cancdence CIS数据库的ODBC连接<br>    第二个为直连CNILG服务器上Access数据库的ODBC连接(AD共用)<br>   第三个为CNILG服务器上Access数据库的文件连接<br>    第四个为CNILX服务器上Access数据库的文件连接<br>所以第一个使用SAPMaxDB连接,后三个使用AccessDB连接。<br>数据库打开状态的Flash信息全改为英文。<br>页面显示全为英文|
| 3.13.0  | 2026-01-24 | 增加DB同步状态自动查看页面并实现功能,使用多线程方式实现，不会跟其它页面任务冲突<br>可以设置:<br>1.检查哪个数据库；<br>2.检查的SAP_Number清单；<br>3. 接收通知邮箱地址; <br>4.检查间隔;<br>5.最大检查次数|
| 3.13.1  | 2026-01-25 | 增加DB sync status check中diff_count=0时结束检查的程序逻辑。|
| 3.14.0  | 2026-01-25 | dbsynccheck页面增加Check Scope Selection功能<br>可以选择仅对比部分列，仅关心eCAD相关的数据<br>程序增加此部分的实现，并且对比结果相同的单元格底色设置为浅绿。|
| 4.0.0  | 2026-01-25 | 主程序中db实例不再使用全局，使用from flask import g方式在每个需要的地方新建实例，以避免多作用多线程操作下出现的冲突问题。|
| 4.0.1  | 2026-01-25 | excel_mapping中的dataildrawing的位置号错误更改|