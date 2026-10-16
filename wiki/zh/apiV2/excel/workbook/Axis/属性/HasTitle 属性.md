---
title: "HasTitle (属性)"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/Axis/%E5%B1%9E%E6%80%A7/HasTitle%20%E5%B1%9E%E6%80%A7.html
updated: 1789700916000
hash: BvBWAGhe
---

# HasTitle (属性)

如果坐标轴或图表有可见标题，则为 **True**。**Boolean** 类型，可读写。

## 说明

坐标轴标题由 **[AxisTitle](./../../AxisTitle/AxisTitle 对象.html)** 对象代表。

## 示例

```javascript
/*此示例为图表工作表 Chart1 中图表的分类坐标轴添加坐标轴标志。*/
function test() {
    let axis = Application.Charts.Item("Chart1").ChartObjects(1).Chart.Axes().Item(xlCategory)
    axis.HasTitle = true
    axis.AxisTitle.Text = "July Sales"
}
```

```javascript
/*此示例将工作表 Sheet1 中第一张图表的数值轴标题设置为不可见。*/
function test() {
    let axis = Application.Sheets.Item("Sheet1").ChartObjects(1).Chart.Axes().Item(xlValue)
    axis.HasTitle = false
}
```
