---
title: "DCount 方法"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/WorksheetFunction/%E6%96%B9%E6%B3%95/DCount%20%E6%96%B9%E6%B3%95.html
---
# DCount (方法)

计算列表或数据库的列中符合指定条件并且包含数字的单元格的个数。

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
/*本示例为C1等单元格分别赋值，并使用DCount方法计算A3:B6相应的列中符合指定条件并且包含数字的单元格的个数。*/
function test() {
    Range("C1").Value2 = "商品"
    Range("D1").Value2 = "价格"
    Range("C2").Value2 = "篮球"
    Range("A3").Value2 = "商品"
    Range("A4").Value2 = "篮球"
    Range("A5").Value2 = "足球"
    Range("A6").Value2 = "篮球"
    Range("D2").Value2 = ">100"
    Range("B3").Value2 = "价格"
    Range("B4").Value2 = 200
    Range("B5").Value2 = 100
    Range("B6").Value2 = 250
    let dcount1 = Application.WorksheetFunction.DCount(Range("A3:B6"), Range("D1").Value2, Range("C1:D2"))
    console.log(dcount1)
}
```

```javascript
/*本示例为A1等单元格分别赋值，使用DCount方法计算C3:D6相应的列中符合指定条件并且包含数字的单元格的个数，并将结果赋值于E2单元格。*/
function test() {
    Range("A1").Value2 = "性别"
    Range("B1").Value2 = "身高"
    Range("A2").Value2 = "男"
    Range("C3").Value2 = "性别"
    Range("C4").Value2 = "男"
    Range("C5").Value2 = "男"
    Range("C6").Value2 = "女"
    Range("B2").Value2 = ">175"
    Range("D3").Value2 = "身高"
    Range("D4").Value2 = 180
    Range("D5").Value2 = 174
    Range("D6").Value2 = 165
    Range("E2").Value2 = Application.WorksheetFunction.DCount(Range("C3:D6"), Range("B1").Value2, Range("A1:B2"))
}
```
