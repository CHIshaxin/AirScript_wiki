---
title: "LegendKey (属性)"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/LegendEntry/%E5%B1%9E%E6%80%A7/LegendKey%20%E5%B1%9E%E6%80%A7.html
updated: 1789700916000
hash: CME4QTAV
---

# LegendKey (属性)

返回一个 **[LegendKey](./../../LegendKey/LegendKey 对象.html)** 对象，该对象表示与图例项相关的图例标示。

## 示例

```javascript
/*本示例显示 Chart1 上图例项一的图例标示的宽度。*/
function test() {
    console.log(Application.Charts.Item("Chart1").ChartObjects(1).Chart.Legend.LegendEntries(1).LegendKey.Width)
}
```

```javascript
/*下例显示 Sheet1 中的第一个图表的第二个图例项的图例标示的高度。*/
function test() {
    let legend = Application.Worksheets.Item("Sheet1").ChartObjects(1).Chart.Legend
    console.log(legend.LegendEntries(2).LegendKey.Height)
}
```
