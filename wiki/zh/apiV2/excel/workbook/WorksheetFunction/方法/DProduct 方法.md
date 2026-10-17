---
title: "DProduct (方法)"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/WorksheetFunction/%E6%96%B9%E6%B3%95/DProduct%20%E6%96%B9%E6%B3%95.html
updated: 1789700916000
hash: KveqayGO
---

# DProduct (方法)

计算列表或数据库的列中符合指定条件的值的乘积。

## 说明

- 由于在单元格中键入文本或值时等号用来指示一个公式，因此 ET 将计算键入内容；但是，这可能会导致产生意外的筛选结果。要为文本或值指示等号比较运算符，请在条件区域的相应单元格中将条件作为字符串表达式键入： **="=entry"** 其中，entry 是要查找的文本或值。例如： | 在单元格中键入的内容 | ET 计算和显示的内容 | | ------------ | -------------- | | ="=Davolio" | =Davolio | | ="=3000" | =3000 |
- ET 在筛选文本数据时不区分大小写。但是，可使用公式执行区分大小写的搜索。

## 参数

| 属性 | 数据类型 | 必填 | 说明 |
| --- | --- | --- | --- |
| Arg1 | Range | 必填 | Database 为构成列表或数据库的单元格区域。数据库是相关数据的列表，其中包含相关信息的行为记录，而包含数据的列为字段。列表的第一行包含每一列的标签。 |
| Arg2 | any | 必填 | Field 指示函数所使用的列。输入使用双引号括起来的列标签，如“使用年数”或“产量”，或者输入代表列表中列位置的数字（不带引号）：1 代表第一列，2 代表第二列，依此类推。 |
| Arg3 | any | 必填 | Criteria 为包含指定条件的单元格区域。只要参数 criteria 至少包含一个列标签和列标签下方用于为该列指定条件的一个单元格，就可以为它使用任意区域。 |

## 返回值

Double

## 示例

```javascript
/*本示例为C1等单元格分别赋值，并使用DProduct方法计算从C1:D4相应的列中提取符合指定条件的值的乘积。*/
function test() {
    Range("C1").Value2 = "水果"
    Range("C2").Value2 = "西瓜"
    Range("C3").Value2 = "荔枝"
    Range("C4").Value2 = "香蕉"
    Range("D1").Value2 = "单价"
    Range("D2").Value2 = 2.3
    Range("D3").Value2 = 10
    Range("D4").Value2 = 3.4
    Range("E1").Value2 = "单价"
    Range("E2").Value2 = ">3"
    Range("F1").Value2 = "单价"
    Range("F2").Value2 = "<30"
    let det1 = Application.WorksheetFunction.DProduct(Range("C1:D4"), Range("D1").Value2, Range("E1:F2"))
    console.log(det1)
}
```

```javascript
/*本示例为C1等单元格分别赋值，并使用DProduct方法计算从C3:E6相应的列中提取符合指定条件的值的乘积。*/
function test() {
    Range("C1").Value2 = "分数"
    Range("C2").Value2 = ">=60"
    Range("D1").Value2 = "总和"
    Range("E1").Value2 = "分数"
    Range("E2").Value2 = "<=100"
    Range("C3").Value2 = "专业"
    Range("C4").Value2 = "计算机"
    Range("C5").Value2 = "计算机"
    Range("C6").Value2 = "计算机"
    Range("D3").Value2 = "分数"
    Range("D4").Value2 = 56.5
    Range("D5").Value2 = 61
    Range("D6").Value2 = 80
    Range("E3").Value2 = "姓名"
    Range("E4").Value2 = "小李"
    Range("E5").Value2 = "小龙"
    Range("E6").Value2 = "张三"
    Range("D2").Value2 = Application.WorksheetFunction.DProduct(Range("C3:E6"), 2, Range("C1:E2"))
}
```
