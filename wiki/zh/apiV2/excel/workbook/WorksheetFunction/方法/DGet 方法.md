---
title: "DGet (方法)"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/WorksheetFunction/%E6%96%B9%E6%B3%95/DGet%20%E6%96%B9%E6%B3%95.html
updated: 1789700916000
hash: B0fk8nqY
---

# DGet (方法)

从列表或数据库的列中提取符合指定条件的单个值。

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

Variant

## 示例

```javascript
/*本示例为D1等单元格分别赋值，并使用DGet方法从A3:C6相应的列中提取符合指定条件的单个值。*/
function test() {
    Range("D1").Value2 = "树种"
    Range("D2").Value2 = "苹果树"
    Range("E1").Value2 = "高度"
    Range("E2").Value2 = ">10"
    Range("F1").Value2 = "产量"
    Range("G1").Value2 = "高度"
    Range("G2").Value2 = "<16"
    Range("A3").Value2 = "树种"
    Range("A4").Value2 = "苹果树"
    Range("A5").Value2 = "梨树"
    Range("A6").Value2 = "苹果树"
    Range("B3").Value2 = "高度"
    Range("B4").Value2 = 18
    Range("B5").Value2 = 1.2
    Range("B6").Value2 = 14
    Range("C3").Value2 = "产量"
    Range("C4").Value2 = 15
    Range("C5").Value2 = 20
    Range("C6").Value2 = 25
    let det1 = Application.WorksheetFunction.DGet(Range("A3:C6"), Range("E1").Value2, Range("D1:G2"))
    console.log(det1)
}
```
