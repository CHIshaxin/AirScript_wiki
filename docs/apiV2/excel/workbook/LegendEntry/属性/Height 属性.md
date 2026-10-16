---
title: "Height 属性"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/LegendEntry/%E5%B1%9E%E6%80%A7/Height%20%E5%B1%9E%E6%80%A7.html
---
# Height (属性)

返回一个 **Double** 值，它代表对象的高度（以磅为单位）。

## 示例

```javascript
/*下例显示 Sheet1 中的第一个图表的第二个图例项的高度。*/
function test() {
    let legend = Application.Worksheets.Item("Sheet1").ChartObjects(1).Chart.Legend
    console.log(legend.LegendEntries().Item(2).Height)
}
```

```javascript
/*下例显示 Chart1 上第一个图例项的高度。*/
function test() {
    let legend = Application.Charts.Item("Chart1").ChartObjects(1).Chart.Legend
    console.log(legend.LegendEntries().Item(1).Height)
}
```
