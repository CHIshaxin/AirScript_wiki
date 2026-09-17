---
title: "Select 方法"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/AxisTitle/%E6%96%B9%E6%B3%95/Select%20%E6%96%B9%E6%B3%95.html
---
# Select (方法)

选择对象。

## 返回值

一个代表所选对象的 Variant 值

## 示例

```javascript
/*本示例选中图表工作表 Chart1 中图表的分类轴标题。*/
function test() {
    let axistitle = Application.Charts.Item("Chart1").ChartObjects(1).Chart.Axes(xlCategory).AxisTitle
    axistitle.Select()
}
```

```javascript
/*本示例将工作表 Sheet1 中第一张图表的数值轴标题设置为可见，并选中该标题。*/
function test() {
    let axis = Application.Sheets.Item("Sheet1").ChartObjects(1).Chart.Axes(xlValue)
    axis.HasTitle = true
    axis.AxisTitle.Select()
}
```
