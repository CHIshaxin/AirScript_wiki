---
title: "FormulaR1C1Local 属性"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/AxisTitle/%E5%B1%9E%E6%80%A7/FormulaR1C1Local%20%E5%B1%9E%E6%80%A7.html
---
# FormulaR1C1Local (属性)

获取或设置一个 **String** 值，该值以用户的语言使用 R1C1 样式表示法来表示对象的公式。可读/写。

## 示例

```javascript
/*本示例将图表工作表 Chart1 中图表的数值轴标题设置为工作表 Sheet1 中 A7 单元格的内容。*/
function test() {
    let axistitle = Application.Charts.Item("Chart1").ChartObjects(1).Chart.Axes(xlValue).AxisTitle
    axistitle.FormulaR1C1Local = "='Sheet1'!R7C1"
}
```

```javascript
/*本示例将工作表 Sheet1 中第一张图表的分类轴标题设置为该工作表 C10 单元格的内容。*/
function test() {
    let axistitle = Application.Sheets.Item("Sheet1").ChartObjects(1).Chart.Axes(xlCategory).AxisTitle
    axistitle.FormulaR1C1Local = "='Sheet1'!R10C3"
}
```
