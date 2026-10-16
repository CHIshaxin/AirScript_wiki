---
title: "FormulaR1C1 (属性)"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/AxisTitle/%E5%B1%9E%E6%80%A7/FormulaR1C1%20%E5%B1%9E%E6%80%A7.html
updated: 1789700916000
hash: CUMU7T0P
---

# FormulaR1C1 (属性)

获取或设置一个 **String** 值，该值以英语使用 R1C1 样式表示法来表示对象的公式。可读/写。

## 示例

```javascript
/*本示例将图表工作表 Chart1 中图表的数值轴标题设置为工作表 Sheet1 中 F2 单元格的内容。*/
function test() {
    let axistitle = Application.Charts.Item("Chart1").ChartObjects(1).Chart.Axes(xlValue).AxisTitle
    axistitle.FormulaR1C1 = "='Sheet1'!R2C6"
}
```

```javascript
/*本示例判断如果工作表 Sheet1 中第一张图表的分类轴标题设置为可见，则将该分类轴标题设置为该工作表 B10 单元格的内容。*/
function test() {
    let axes = Application.Worksheets.Item("Sheet1").ChartObjects(1).Chart.Axes(xlCategory)
    if (axes.HasTitle) {
        axes.AxisTitle.FormulaR1C1 = "='Sheet1'!R10C2"
    }
}
```
