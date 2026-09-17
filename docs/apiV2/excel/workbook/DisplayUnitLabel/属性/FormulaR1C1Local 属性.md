---
title: "FormulaR1C1Local 属性"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/DisplayUnitLabel/%E5%B1%9E%E6%80%A7/FormulaR1C1Local%20%E5%B1%9E%E6%80%A7.html
---
# FormulaR1C1Local (属性)

获取或设置一个 **String** 值，该值以用户的语言使用 R1C1 样式表示法来表示对象的公式。可读/写。

## 返回值

String

## 示例

```javascript
/*本示例将图表工作表 Chart1 中数值轴上的单位标签设置为工作表 Sheet1 中 D5 单元格的内容。*/
function test() {
    let displayUnitLabel = Application.Charts.Item("Chart1").ChartObjects(1).Chart.Axes(xlValue).DisplayUnitLabel
    displayUnitLabel.FormulaR1C1Local = "='Sheet1'!R5C4"
}
```

```javascript
/*下例显示以用户的语言使用 R1C1 样式表示法来表示 Sheet1 上的第一个图表数值轴上的单位标签的公式。*/
function test() {
    let displayUnitLabel = Application.Worksheets.Item("Sheet1").ChartObjects(1).Chart.Axes(xlValue).DisplayUnitLabel
    console.log(displayUnitLabel.FormulaR1C1Local)
}
```
