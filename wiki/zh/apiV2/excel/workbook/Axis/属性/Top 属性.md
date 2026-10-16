---
title: "Top (属性)"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/Axis/%E5%B1%9E%E6%80%A7/Top%20%E5%B1%9E%E6%80%A7.html
updated: 1789700916000
hash: D0P_3e5O
---

# Top (属性)

返回一个 **Double** 值，它代表从对象的上边缘到工作表第一行顶部或图表上的图表区顶部的距离（以磅为单位）。

## 示例

```javascript
/*本示例显示图表工作表 Chart1 中图表的分类轴到图表顶部的距离。*/
function test() {
    let axis = Application.Charts.Item("Chart1").ChartObjects(1).Chart.Axes(xlCategory)
    console.log(`分类轴到图表顶部的距离：${axis.Top} 磅`)
}
```

```javascript
/*本示例显示工作表 Sheet1 中第一张图表的数值轴到图表顶部的距离是否为 0 磅。*/
function test() {
    let axis = Application.Sheets.Item("Sheet1").ChartObjects(1).Chart.Axes(xlValue)
    console.log(axis.Top == 0)
}
```
