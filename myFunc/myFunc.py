# -*- coding:utf-8 -*-
# Python3 + xlwings
# Author: https://github.com/playGitboy/AdvancedExcelFunc

import xlwings as xw
import numpy as np
import pandas as pd
from xlwings.constants import RemoveDocInfoType, Constants
from collections import Counter

@xw.sub
def main():
    """演示：一键修改当前工作表数据格式（xlwings 0.16.0+）"""
    wb = xw.Book.caller()
    # 删除文档属性和个人信息
    wb.api.RemoveDocumentInformation(RemoveDocInfoType.xlRDIDocumentProperties)
    # 修正标题栏显示；日期、金额等单元格格式
    sheets0 = wb.sheets[0]
    rangeFirstRow = sheets0.range("A1").expand("right")
    for i in rangeFirstRow:
        col = i.api.EntireColumn
        if i.value.find("期") != -1:
            col.NumberFormat = "yyyy-mm-dd"
        elif i.value.find("价") != -1 or i.value.find("额") != -1:
            col.NumberFormat = "0.00"
            col.HorizontalAlignment = Constants.xlRight
        else:
            col.NumberFormat = "@"
    rangeFirstRow.api.HorizontalAlignment = Constants.xlCenter
    rangeFirstRow.api.Font.Bold = True
    # 设置列宽、刷新数据等
    sheets0.api.UsedRange.value = sheets0.api.UsedRange.value
    sheets0.api.UsedRange.VerticalAlignment = Constants.xlCenter
    sheets0.api.UsedRange.Font.Size = 10
    sheets0.autofit()
    # 控制过大列宽
    for i in rangeFirstRow:
        if i.api.EntireColumn.ColumnWidth > 30:
            i.api.EntireColumn.ColumnWidth = 30


def _l2dTranpose(llSrc):
    """二维列表转置"""
    return [[y[x] for y in llSrc] for x in range(len(llSrc[0]))]


def _d2dCounter(llSrc):
    """
        列表元素计数
        [1,2,3,1,3] → Counter({'1': 2, '2': 1, '3': 2})
    """
    return Counter([str(i) for i in llSrc])


def _fmtArg(x):
    """参数格式化：将所有类型参数统一规整为0或1"""
    try:
        x = 1 if x > 0 else 0
    except:
        x = 0
    return x

def _findNearest(llSrc,val,match_mode):
    """近似匹配：从llSrc列表中查找大于/小于val近似值"""
    lTmp = []
    for i in llSrc:
        try:
            lTmp.append(float(i[0]))
        except:
            lTmp.append(None)
    if match_mode>0:
        minVal = min(filter(None.__ne__,[i for i in lTmp if i==None or i>=val[0]]))
    else:
        minVal = max(filter(None.__ne__,[i for i in lTmp if i==None or i<=val[0]]))
    return lTmp.index(minVal)

def _getVals(lDst,lSrc,val):
    """列表花式索引：从lSrc列表中查找所有val的index，返回lDst对应位置的值"""
    if len(lDst) == len(lSrc):
        idxs = [i for i,v in enumerate(lSrc) if v == val]
        llResult = [lDst[idx] for idx in idxs]
        lTmp = []
        for i in llResult:
            lTmp.extend(i)
        return [lTmp]
    else:
        return None

@xw.func
@xw.arg("lookup_value", doc="查找值")
@xw.arg("lookup_array", doc="查找值所处数据区域")
@xw.arg("return_array", doc="返回值所处数据区域")
@xw.arg("search_mode", doc="0返回最后一个匹配值 1返回第一个匹配值(默认)")
@xw.ret(expand="table")
def myXLOOKUP(lookup_value, lookup_array, return_array, search_mode=1):
    """数据查找"""
    if not isinstance(lookup_value, list) and isinstance(lookup_array, list) and isinstance(return_array, list) and len(lookup_array) == len(return_array):
        if search_mode == 1:
            return return_array[lookup_array.index(lookup_value)]
        else:
            return return_array[::-1][lookup_array[::-1].index(lookup_value)]
    else:
        return "参数错误！"

@xw.func
@xw.arg("lookup_value", ndim=2, doc="查找值")
@xw.arg("lookup_array", ndim=2, doc="查找值所处数据区域")
@xw.arg("return_array", ndim=2, doc="返回值所处数据区域")
@xw.arg("match_mode", doc="0精确匹配(默认) -1匹配邻近较小值 1匹配邻近较大值")
@xw.arg("search_mode", doc="0返回全部匹配值 -1返回最后一个匹配值 1返回首个匹配值(默认)")
@xw.ret(expand="table")
def myXLOOKUP2(lookup_value, lookup_array, return_array, match_mode =0, search_mode =1):
    """数据查找"""
    try:
        if len(lookup_array) == len(return_array):
            # 若lookup_value为区域，则逐个元素递归调用本函数对应取值
            isRangeSearch = len(lookup_value[0]) > 1 and len(lookup_array[0]) > 1
            if not isRangeSearch:
                if len(lookup_value) > 1 or len(lookup_value[0]) > 1:
                    if search_mode != 0:
                        lookup_value = [[i] for ii in lookup_value for i in ii]
                        return [myXLOOKUP2([v], lookup_array, return_array, match_mode, search_mode) for v in lookup_value]
                    else:
                        return "批量查询模式search_mode不能为0！"
            # 反查模式
            if search_mode == -1:
                # 横/纵向查找
                if not isRangeSearch and len(lookup_array[0]) > 1:
                    lookup_array = [lookup_array[0][::-1]]
                    return_array = [return_array[0][::-1]]
                else:
                    lookup_array = lookup_array[::-1]
                    return_array = return_array[::-1]
            # 精确匹配和近似匹配算法不同
            if match_mode == 0:
                if search_mode == 0:
                    return _getVals(return_array, lookup_array, lookup_value[0])
                else:
                    if not isRangeSearch and len(lookup_array[0]) > 1:
                        return return_array[0][lookup_array[0].index(lookup_value[0][0])]
                    else:
                        return return_array[lookup_array.index(lookup_value[0])]
            return return_array[_findNearest(lookup_array,lookup_value[0],match_mode)]
        else:
            return "数组公式请使用Ctrl+Shift+Enter录入！"
    except Exception as e:
        return "请检查参数！" if type(e) == ValueError else e.args


@xw.func
@xw.arg("rows", doc="行数")
@xw.arg("columns", doc="列数(默认1)")
@xw.arg("start", doc="起始值(默认1)")
@xw.arg("step", doc="步进值(默认1)")
@xw.ret(expand="table")
def mySEQUENCE(rows, columns=1, start=1, step=1):
    """生成等差序列数"""
    rows, columns, start, step = map(int, [rows, columns, start, step])
    return np.arange(start, rows*columns*step+start, step).reshape(rows, columns)


@xw.func
@xw.arg("data_array", pd.DataFrame, index=False, header=False, doc="数据区域")
@xw.arg("sort_index", doc="按第几列排序(默认1)")
@xw.arg("sort_order", doc="0降序 1升序(默认)")
@xw.arg("have_header", doc="0不包含标题 1包含标题(默认)")
@xw.arg("axis", doc="0按行排序(默认) 1按列排序")
@xw.ret(index=False, header=False, expand="table")
def mySORT(data_array, sort_index=1, sort_order=1, have_header=1, axis=0):
    """数据排序"""
    try:
        have_header, sort_order, axis = map(_fmtArg, [have_header, sort_order, axis])
        if axis:
            data_array = data_array.T
        if sort_index > data_array.shape[1] or sort_index <= 0:
            return "排序列索引值超出区域范围！"
        if have_header:
            data_array.columns = data_array.loc[0]
            data_array = data_array.drop(0)
        else:
            data_array.columns = range(data_array.shape[1])
        df = data_array.sort_values(by=data_array.columns[sort_index - 1],
                                    ascending=sort_order)
        if have_header:
            df2 = df.copy().drop(index=df.index)
            df2.loc[0] = df2.columns
            df = df2.append(df, ignore_index=True)
        return df.T if axis else df
    except TypeError:
        return "排序列数据类型不一致，选区是否包含标题？"


@xw.func
@xw.arg("data_array", np.array, doc="数据区域")
@xw.arg("include", np.array, doc="过滤条件")
@xw.arg("if_empty", doc="无数据返回值(默认空)")
@xw.ret(expand="table")
def myFILTER(data_array, include, if_empty=""):
    """数据筛选"""
    include = include.astype(bool)
    return data_array[include] if len(data_array[include]) else if_empty


@xw.func
@xw.arg("delimiter", doc="分隔符")
@xw.arg("ignore_empty", doc="空白值 0保留 1忽略")
@xw.arg('datas', ndim=2, doc="待合并区域(支持多个区域)")
def myTEXTJOIN(delimiter, ignore_empty, *datas):
    """分隔符合并区域文本"""
    lResult = []
    for llData in datas:
        lStr = [str(i)[:-2] if str(i).endswith(".0") else str(i) for lData in llData for i in lData]
        if ignore_empty:
            lResult.extend([i for i in lStr if i != "None" and i != ""])
        else:
            lResult.extend([i.replace("None", "") for i in lStr])
    return delimiter.join(lResult)


@xw.func
@xw.arg("rows", doc="行数")
@xw.arg("columns", doc="列数")
@xw.arg("iMin", doc="最小值(默认无)")
@xw.arg("iMax", doc="最大值(默认无)")
@xw.arg("bInt", doc="0返回小数 1返回整数(设置取数范围时生效，默认为1)")
@xw.ret(expand="table")
def myRANDARRAY(rows, columns, iMin=None, iMax=None, bInt=1):
    """生成随机数矩阵"""
    rows, columns = map(int, [rows, columns])
    if iMin != None and iMax != None:
        if bInt:
            return np.random.randint(iMin, iMax, (rows, columns))
        else:
            return iMin + (iMax-iMin)*np.random.rand(rows, columns)
    return np.random.rand(rows, columns)


@xw.func
@xw.arg("data_array", ndim=2, doc="数据区域")
@xw.arg("axis", doc="0按行筛选(默认) 1按列筛选")
@xw.arg("exactly_once", doc="0提取全部不重复数据 1仅提取出现一次的数据(默认0)")
@xw.ret(expand="table")
def myUNIQUE(data_array, axis=0, exactly_once=0):
    """筛选不重复数据"""
    if axis:
        list2 = _l2dTranpose(data_array)
        if exactly_once:
            dTemp = _d2dCounter(list2)
            return _l2dTranpose([eval(i) for i in dTemp.keys() if dTemp[i] == 1])
        else:
            return _l2dTranpose(sorted([list(t) for t in set(tuple(_) for _ in list2)], key=list2.index))
    else:
        if exactly_once:
            dTemp = _d2dCounter(data_array)
            return [eval(i) for i in dTemp.keys() if dTemp[i] == 1]
        else:
            return sorted([list(t) for t in set(tuple(_) for _ in data_array)], key=data_array.index)


@xw.func
@xw.arg("data_array", xw.Range, doc="数据区域")
# range为区域则formula返回二维嵌套元组；range为单元格，返回字符串
def mySUMVALUE(data_array):
    """忽略公式/文本，汇总区域数字（包括文本型数字）"""
    rangeTuple = data_array.formula
    if isinstance(rangeTuple, tuple):
        lValues = [i for ii in rangeTuple for i in ii if not i.startswith("=")]
        fTotal = 0
        for i in lValues:
            try:
                fTotal += float(i)
            except:
                continue
    return fTotal


@xw.func
@xw.arg("rows", doc="行数")
@xw.arg("columns", doc="列数")
@xw.arg("szType", doc="name/ssn/address/company/email/phone_number...")
@xw.ret(expand="table")
def myFAKER(rows, columns=1, szType="name"):
    """随机生成测试数据"""
    try:
        from faker import Faker
        f = Faker("zh_CN")
        rows, columns = map(int, [rows, columns])
        lRand = np.array([getattr(f,szType)() for i in range(rows * columns)])
        return lRand.reshape(rows, columns)
    except ImportError:
        return "faker未安装，请运行：pip install faker"

@xw.func
@xw.arg("data_array", np.array, doc="数据区域")
@xw.arg("rows", np.array, doc="行数")
@xw.arg("columns", np.array, doc="列数")
@xw.arg("axis", doc="0纵轴 1横轴(默认)")
@xw.ret(expand="table")
def myNDimension(data_array, rows=0, columns=0, axis=1):
    """按指定维度转换数据区域"""
    if not rows * columns:
        rows,columns = map(int,[rows,columns])
        if rows:
            columns =  int(__import__("math").ceil(data_array.size / rows))
        else:
            rows =  int(__import__("math").ceil(data_array.size / columns))
        if axis == 0:
            data_array = data_array.T
        data_array = np.append(data_array.astype(str), int(rows * columns - data_array.size) * [""])
        data_array[data_array == "nan"] = ""
        data_array[np.char.endswith(data_array,".0")] = np.char.replace(data_array[np.char.endswith(data_array,".0")],".0","")
        return data_array.reshape(rows, columns)
    return "行数/列数二选一，另一个会自动生成！"


@xw.func
@xw.arg("dim2_range", pd.DataFrame, index=False, doc="二维表数据区域（含首行/首列的标题）")
@xw.ret(index=False, expand="table")
def mySTACK(dim2_range, left_index=1, top_header=1):
    """二维表转一维表"""
    left_index,top_header = map(_fmtArg,[left_index,top_header])
    if left_index * top_header:
        return dim2_range.set_index([dim2_range.columns[0]]).stack().reset_index()
    elif left_index==0:
        return dim2_range.stack().reset_index().iloc[:,1:]
    elif top_header==0:
        df = dim2_range.T.reset_index()
        df.columns = df.loc[0]
        df = df.drop(0).stack().to_frame().rename(columns = {0:"DATA"})
        df.index.names = ["TMP","INDEX"]
        return df.reset_index(level="INDEX")


if __name__ == "__main__":
    xw.serve()
    xw.books.active.set_mock_caller()
    main()
