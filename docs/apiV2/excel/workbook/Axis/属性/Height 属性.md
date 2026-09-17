---
title: "Height 属性"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/Axis/%E5%B1%9E%E6%80%A7/Height%20%E5%B1%9E%E6%80%A7.html
---
# Height (属性)

返回一个 **Double** 值，它代表对象的高度（以磅为单位）。

## 示例

```javascript
/*本示例显示工作表 Sheet1 中第一张图表的数值坐标轴的高度。*/
function test() {
    let axis = Application.Sheets.Item("Sheet1").ChartObjects(1).Chart.Axes(xlValue)
    console.log(axis.Height)
}
```

```javascript
/*本示例显示图表工作表 Chart1 中图表的分类坐标轴的高度。*/
function test() {
    let axis = Application.Charts.Item("Chart1").ChartObjects(1).Chart.Axes(xlCategory)
    console.log(axis.Height)
}
```
