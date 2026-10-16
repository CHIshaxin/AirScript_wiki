---
title: "Left (属性)"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/Axis/%E5%B1%9E%E6%80%A7/Left%20%E5%B1%9E%E6%80%A7.html
updated: 1789700916000
hash: CW7OI4SJ
---

# Left (属性)

返回一个 **Double** 值，它代表从对象左边缘到图表区左边缘的距离（以磅为单位）。

## 示例

```javascript
/*本示例显示图表工作表 Chart1 中图表的分类坐标轴到图表区左边缘的距离。*/
function test() {
    let axis = Application.Charts.Item("Chart1").ChartObjects(1).Chart.Axes(xlCategory)
    console.log(axis.Left)
}
```

```javascript
/*本示例显示工作表 Sheet1 中第一张图表的数值坐标轴到图表区左边缘的距离。*/
function test() {
    let axis = Application.Sheets.Item("Sheet1").ChartObjects(1).Chart.Axes(xlValue)
    console.log(axis.Left)
}
```
