---
title: "Formula 属性"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/AxisTitle/%E5%B1%9E%E6%80%A7/Formula%20%E5%B1%9E%E6%80%A7.html
---
# Formula (属性)

获取或设置一个 **String** 值，该值以英语使用 A1 样式表示法来表示对象的公式。可读/写。

## 说明

有关 A1 样式表示法的详细信息，请参阅操作方法：使用 A1 表示法引用单元格和区域。

## 示例

```javascript
/*本示例设置图表工作表 Chart1 中图表的分类轴标题为工作表 Sheet1 中 E1 单元格的内容。*/
function test() {
    let axistitle = Application.Charts.Item("Chart1").ChartObjects(1).Chart.Axes(xlCategory).AxisTitle
    axistitle.Formula = "=Sheet1!$E$1"
}
```

```javascript
/*本示例将工作表 Sheet1 中第一张图表的数值轴标题设置为可见，并将该数值轴标题设置为该工作表 E5 单元格的内容。*/
function test() {
    let axes = Application.Worksheets.Item("Sheet1").ChartObjects(1).Chart.Axes(xlValue)
    axes.HasTitle = true
    axes.AxisTitle.Formula = "=Sheet1!$E$5"
}
```
