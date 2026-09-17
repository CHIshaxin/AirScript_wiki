---
title: "ReversePlotOrder 属性"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/Axis/%E5%B1%9E%E6%80%A7/ReversePlotOrder%20%E5%B1%9E%E6%80%A7.html
---
# ReversePlotOrder (属性)

如果 ET 的绘图区数据点的顺序为从后往前，则该属性值为 **True**。**Boolean** 类型，可读写。

## 说明

该属性不能用于雷达图。

## 示例

```javascript
/*本示例将图表工作表 Chart1 数值轴的绘图区数据点的顺序设置为从后往前。*/
function test() {
    let axis = Application.Charts.Item("Chart1").ChartObjects(1).Chart.Axes(xlValue)
    axis.ReversePlotOrder = true
}
```

```javascript
/*本示例判断工作表 Sheet1 中第一张图表的分类轴的绘图区数据点的顺序，并通知用户。*/
function test() {
    let axis = Application.Sheets.Item("Sheet1").ChartObjects(1).Chart.Axes(xlCategory)
    if (axis.ReversePlotOrder) {
        console.log("分类轴的绘图区数据点的顺序为从后往前")
    } else {
        console.log("分类轴的绘图区数据点的顺序为从前往后")
    }
}
```
