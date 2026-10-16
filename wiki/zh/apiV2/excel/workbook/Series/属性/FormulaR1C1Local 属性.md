---
title: "FormulaR1C1Local (属性)"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/Series/%E5%B1%9E%E6%80%A7/FormulaR1C1Local%20%E5%B1%9E%E6%80%A7.html
updated: 1789700916000
hash: YNxPgQfu
---

# FormulaR1C1Local (属性)

返回或设置指定对象的公式，使用用户语言 R1C1 格式符号表示。**String** 型，可读写。

## 说明

如果指定单元格包含常量，此属性返回的就是该常量。如果单元格为空，此属性返回一个空字符串。如果单元格包含公式，此属性将该公式作为字符串返回，所用格式与在编辑栏（包括等号）中显示时的格式相同。

如果将单元格的格式的值或公式设为日期类型，ET 将检查该单元格的格式是否符合某个日期或时间数组格式，如果不符合，将采用默认的短日期数字格式。

如果指定区域是一维或二维区域，则可将公式指定为 JavaScript 中相同维数的数组。同样，也可在 JavaScript 数组中使用公式。

对多重单元格区域设置公式，则该区域中所有单元格都用此公式填充。

## 示例

```javascript
/*此示例使用用户语言的 R1C1 样式表示法设置第一张工作表中第一个图表对象的第一个系列。*/
function test() {
    let series1 = Application.Worksheets.Item(1).ChartObjects(1).Chart.SeriesCollection(1)
    series1.FormulaR1C1Local = "=SERIES(Sheet1!R1C5,Sheet1!R2C1:R2C4,Sheet1!R1C1:R1C4,1)"
}
```

```javascript
/*此示例使用用户语言的 R1C1 样式表示法将一个公式赋值给图表工作表 Chart1 上的第二个数据系列。*/
function test() {
    let series2 = Application.Charts.Item("Chart1").ChartObjects(1).Chart.SeriesCollection(2)
    series2.FormulaR1C1Local = "=SERIES(Sheet1!R1C5,Sheet1!R2C1:R2C4,Sheet1!R1C1:R1C4,1)"
}
```
