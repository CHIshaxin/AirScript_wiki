---
title: "DVar 方法"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/WorksheetFunction/%E6%96%B9%E6%B3%95/DVar%20%E6%96%B9%E6%B3%95.html
---
# DVar (方法)

将列表或数据库的列中符合指定条件的数字作为一个样本，估算总体的方差。

## 说明

- 由于在单元格中键入文本或值时等号用来指示一个公式，因此 ET 将计算键入内容；但是，这可能会导致产生意外的筛选结果。要为文本或值指示等号比较运算符，请在条件区域的相应单元格中将条件作为字符串表达式键入： **="=entry"** 其中，entry 是要查找的文本或值。例如：

| 在单元格中键入的内容 | ET 计算和显示的内容 |
| --- | --- |
| ="=Davolio" | =Davolio |
| ="=3000" | =3000 |

- ET 在筛选文本数据时不区分大小写。但是，可使用公式执行区分大小写的搜索。

## 参数

| 属性 | 数据类型 | 必填 | 说明 |
| --- | --- | --- | --- |
| Arg1 | Range | 必填 | Database - 构成列表或数据库的单元格区域。数据库是相关数据的列表，其中包含相关信息的行为记录，而包含数据的列为字段。列表的第一行包含每一列的标签。 |
| Arg2 | any | 必填 | Field - 指示函数所使用的列。输入使用双引号括起来的列标签，如“使用年数”或“产量”，或者输入代表列表中列位置的数字（不带引号）：1 代表第一列，2 代表第二列，依此类推。 |
| Arg3 | any | 必填 | Criteria - 包含指定条件的单元格区域。只要参数 criteria 至少包含一个列标签和列标签下方用于为该列指定条件的一个单元格，就可以为它使用任意区域。 |

## 返回值

Double

## 示例

```javascript
/*本示例为A1等单元格分别赋值，并使用DVar方法从A1:C4相应的列中提取符合指定条件的数字作为一个样本，估算总体的方差。*/
function test() {
    Range("A1").Value2 = "树种"
    Range("A2").Value2 = "苹果树"
    Range("A3").Value2 = "桃子树"
    Range("A4").Value2 = "樱花树"
    Range("B1").Value2 = "高度"
    Range("B2").Value2 = 10
    Range("B3").Value2 = 10
    Range("B4").Value2 = 10
    Range("C1").Value2 = "产量"
    Range("C2").Value2 = 17
    Range("C3").Value2 = 20
    Range("C4").Value2 = 21
    Range("E1").Value2 = "产量"
    Range("E2").Value2 = ">10"
    Range("F1").Value2 = "产量"
    Range("F2").Value2 = "<30"
    let det1 = Application.WorksheetFunction.DVar(Range("A1:C4"), Range("C1").Value2, Range("E1:F2"))
    console.log(det1)
}
```

```javascript
/*本示例为A1等单元格分别赋值，使用DVar方法从E3:G6相应的列中提取符合指定条件的数字作为一个样本，估算总体的方差并将结果赋值于B2单元格。*/
function test() {
    Range("A1").Value2 = "身高"
    Range("A2").Value2 = ">145"
    Range("B1").Value2 = "方差"
    Range("C1").Value2 = "身高"
    Range("C2").Value2 = "<190"
    Range("E3").Value2 = "班级"
    Range("E4").Value2 = "二班"
    Range("E5").Value2 = "二班"
    Range("E6").Value2 = "二班"
    Range("F3").Value2 = "身高"
    Range("F4").Value2 = 177.2
    Range("F5").Value2 = 153
    Range("F6").Value2 = 185
    Range("G3").Value2 = "姓名"
    Range("G4").Value2 = "王五"
    Range("G5").Value2 = "李四"
    Range("G6").Value2 = "张三"
    Range("B2").Value2 = Application.WorksheetFunction.DVar(Range("E3:G6"), 2, Range("A1:C2"))
}
```
