---
title: "Width (属性)"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/Axis/%E5%B1%9E%E6%80%A7/Width%20%E5%B1%9E%E6%80%A7.html
updated: 1789700916000
hash: Bg8Y_byz
---

# Width (属性)

返回一个 **Double** 值，它代表对象的宽度（以磅为单位）。

## 示例

```javascript
/*本示例显示工作表 Sheet1 中第一张图表的数值轴的宽度。*/
function test() {
    let axis = Application.Sheets.Item("Sheet1").ChartObjects(1).Chart.Axes(xlValue)
    console.log(axis.Width)
}
```

```javascript
/*本示例显示图表工作表 Chart1 中图表的分类轴的宽度。*/
function test() {
    let axis = Application.Charts.Item("Chart1").ChartObjects(1).Chart.Axes(xlCategory)
    console.log(`分类轴的宽度：${axis.Width} 磅`)
}
```
