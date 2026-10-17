---
title: "Font (属性)"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/TickLabels/%E5%B1%9E%E6%80%A7/Font%20%E5%B1%9E%E6%80%A7.html
updated: 1789700916000
hash: DML2Mrl5
---

# Font (属性)

返回一个 **[Font](./../../Font/Font 对象.html)** 对象，它代表指定对象的字体。

## 示例

```javascript
/*下例将 Sheet1 上的第一个图表分类轴刻度线标签的字体设为加粗。*/
function test() {
    let tickLabels = Application.Worksheets.Item("Sheet1").ChartObjects(1).Chart.Axes(xlCategory).TickLabels
    tickLabels.Font.Bold = true
}
```

```javascript
/*以下示例将 Chart1 上数值轴刻度线标签的字体大小设为 13 磅。*/
function test() {
    let tickLabels = Application.Charts.Item("Chart1").ChartObjects(1).Chart.Axes(xlValue).TickLabels
    tickLabels.Font.Size = 13
}
```
