---
title: "DMin (方法)"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/WorksheetFunction/%E6%96%B9%E6%B3%95/DMin%20%E6%96%B9%E6%B3%95.html
updated: 1789700916000
hash: 9QzFjSkR
---

# DMin (方法)

返回列表或数据库的列中符合指定条件的最小数字。

## 说明

- 由于在单元格中键入文本或值时等号用来指示一个公式，因此 ET 将计算键入内容；但是，这可能会导致产生意外的筛选结果。要为文本或值指示等号比较运算符，请在条件区域的相应单元格中将条件作为字符串表达式键入： **="=entry"** 其中，entry 是要查找的文本或值。例如： | 在单元格中键入的内容 | ET 计算和显示的内容 | | ------------ | -------------- | | ="=Davolio" | =Davolio | | ="=3000" | =3000 |
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
/*本示例为A3等单元格分别赋值，并使用DMin方法从A3:C6相应的列中提取符合指定条件的最小值。*/
function test() {
    Range("A3").Value2 = "树种"
    Range("A4").Value2 = "苹果树"
    Range("A5").Value2 = "梨树"
    Range("A6").Value2 = "香蕉树"
    Range("B3").Value2 = "高度"
    Range("B4").Value2 = 15
    Range("B5").Value2 = 9
    Range("B6").Value2 = 14
    Range("C3").Value2 = "产量"
    Range("C4").Value2 = 15
    Range("C5").Value2 = 20
    Range("C6").Value2 = 25
    Range("E1").Value2 = "高度"
    Range("E2").Value2 = ">10"
    Range("F1").Value2 = "高度"
    Range("F2").Value2 = "<20"
    let det1 = Application.WorksheetFunction.DMin(Range("A3:C6"), Range("B3").Value2, Range("E1:F2"))
    console.log(det1)
}
```

```javascript
/*本示例为B1等单元格分别赋值，使用DMin方法从E3:G6相应的列中提取符合指定条件的最小值，并将结果赋值于C2单元格。*/
function test() {
    Range("B1").Value2 = "分数"
    Range("B2").Value2 = ">90"
    Range("C1").Value2 = "最小值"
    Range("D1").Value2 = "分数"
    Range("D2").Value2 = "<95"
    Range("E3").Value2 = "班级"
    Range("E4").Value2 = "高三二班"
    Range("E5").Value2 = "高三六班"
    Range("E6").Value2 = "高三二班"
    Range("F3").Value2 = "分数"
    Range("F4").Value2 = 95
    Range("F5").Value2 = 93
    Range("F6").Value2 = 94
    Range("G3").Value2 = "姓名"
    Range("G4").Value2 = "王五"
    Range("G5").Value2 = "赵四"
    Range("G6").Value2 = "李三"
    Range("C2").Value2 = Application.WorksheetFunction.DMin(Range("E3:G6"), 2, Range("A1:D2"))
}
```
