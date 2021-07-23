# Excel高阶函数python版  
基于Python的xlwings库实现TEXTJOIN/SEQUENCE/RANDARRAY/UNIQUE/XLOOKUP/SORT/FILTER等7个Excel新增高阶函数，同时新增FAKER(批量生成测试数据，如姓名/用户名/手机/身份证/邮箱等)、STACK、NDimension(按指定维度转换数据区域)、SUMVALUE(忽略公式仅汇总数字)等自定义函数，让普通Excel2007-2013也能用上Excel365新增函数……  

函数功能详见Microsoft帮助，如  
[XLOOKUP](https://support.microsoft.com/zh-cn/office/xlookup-%E5%87%BD%E6%95%B0-b7fd680e-6d10-43e6-84f9-88eae8bf5929)  
[QNIQUE](https://support.microsoft.com/zh-cn/office/unique-%E5%87%BD%E6%95%B0-c5ab87fd-30a3-4ce9-9d1a-40204fb85e1e)  

## 安装和设置  
### 安装  
方法一：下载安装 Anaconda3（推荐）  
方法二：下载安装 Python3.x，然后终端执行  
```  
pip install xlwings  
pip install ipython  
xlwings addin install  
```

相关链接：  
[Anaconda](https://www.anaconda.com/distribution/#download-section)  
[Python](https://www.python.org/downloads/windows/)  
[xlwings UDF](https://docs.xlwings.org/zh_CN/latest/udfs.html)  

### Excel设置(以2007+为例)   
打开任意Excel文件，Alt+F11打开VBE，“工具——引用”勾选xlwings  
Excel“文件——选项——信任中心——信任中心设置——宏设置”勾选“信任对VBA工程对象模型的访问”；为方便测试可勾选“启用所有宏”   

> myFunc.xlsm包含每个自定义函数的功能释义和测试数据  

## xlwings UDF Tips  
UDF仅支持Windows系统（需要Office VBA配合）  
参数、数据默认类型：float/unicode/datetime/None  
Excel单行/列数据读取为python列表：[None, 1.0, 'string']  
Excel的2D区域数据读取为python二维列表：[[None, 1.0, 'a string'], [None, 2.0,'another string']]  

## IDE调试UDF  
VSCode为例，IDE设置步骤： 
> 打开UDF对应的py脚本增加```xw.serve() ```  
> 在待调试的代码行设置断点 
> F5启动调试  

Excel设置步骤:  
> xlwings选项卡选中“Debug UDFs”后点击"Import Functions" 
> 使用待调试函数，IDE即可自动中断代码  

## Excel中xlwings工具使用常见问题  
“Run main”功能使用注意事项  
> 需要xlwings 0.16.0+  
> py脚本名和Excel文件名一致，且不能含特殊字符(如括号、引号等)  
> 支持任意Excel文件(xls/xlsm/xlsx等等)  

重复点击"Import Functions"按钮Excel提示以下错误  
> 弹窗"Could not activate Python COM server, hr = -2147023170 1000"  
> 先点击“Restart UDF Server”再点击“Import Functions”即可 

点击"Import Functions"按钮Excel提示以下错误  
> 弹窗“No command specified in the configuration, cannot autostart server 1000”   
> 勾选了"Debug UDFs"但未检测到后台调试进程——如需调试UDF请查看"IDE调试UDF"注意事项，否则请取消"Debug UDFs"勾选   

> 弹窗“自动化 (Automation) 错误 440”  
> UDF参数名称与python内置函数名称/关键字等冲突，比如min/max/integer等  

> 弹窗“无法使用该函数”  
> UDF函数名与Excel函数名冲突，如"XLOOKUP"编写UDF时建议命名为"myXLOOKUP"以示区别  

> 单元格显示“要求对象”  
> Excel中Alt+F11打开VBE，“工具——引用”勾选xlwings  

> 单元格显示“Could not create Python process. Error message: 拒绝访问”   
> Python被其他程序后台占用，任务管理器结束python进程或重启系统  

## In-Excel SQL  
安装设置xlwings和Excel后，在任意Excel文件中都可以直接使用sql函数查询数据区域  
=sql(SQL Statement, table a, table b, ...)  

注意：  
基于SQLite内存模式，SQL语法与SQLite一致  
不同数据区域按前后顺序默认表别名为a b c d e…  

## 测试环境  
System：Windows 7/10  
Office：2010/2013  
Python：3.7  
xlwings：0.17.1  
(Anaconda3-2020.02)  

## Todo  
各函数主要功能完成，但函数过多未及详测，欢迎提Issues或Fork...

