---
title: "Height (属性)"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/LegendKey/%E5%B1%9E%E6%80%A7/Height%20%E5%B1%9E%E6%80%A7.html
updated: 1789700916000
hash: B5xdbMqm
---

# Height (属性)

返回一个 **Double** 值，它代表对象的高度（以磅为单位）。

## 示例

```javascript
/*下例显示 Sheet1 中的第一个图表的第二个图例项的图例标示的高度。*/
function test() {
    let legend = Application.Worksheets.Item("Sheet1").ChartObjects(1).Chart.Legend
    console.log(legend.LegendEntries(2).LegendKey.Height)
}
```

```javascript
/*本示例显示 Chart1 上第一个图例项的图例标示的高度。*/
function test() {
    console.log(Application.Charts.Item("Chart1").ChartObjects(1).Chart.Legend.LegendEntries(1).LegendKey.Height)
}
```
