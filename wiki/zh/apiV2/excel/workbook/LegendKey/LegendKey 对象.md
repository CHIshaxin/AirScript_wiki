---
title: "LegendKey (对象)"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/LegendKey/LegendKey%20%E5%AF%B9%E8%B1%A1.html
updated: 1789700916000
hash: Dvfi0-Vy
---

# LegendKey (对象)

代表图表图例中的图例标示。

## 说明

每个图例项标示都是一个图形，它将图例项标示以及与之相关联的图表中的系列或趋势线链接起来。图例项标示链接到与之相关联的系列或趋势线的方式为：更改一个的格式会同时更改另一个的格式。

使用 **LegendKey** 属性可返回 **[LegendKey](./../LegendKey/LegendKey 对象.html)** 对象。比如更改图例顶部的图例项标示的标志背景色，这样会同时更改与该图例项相关的系列中每一个数据点的格式。相关的系列必须支持数据标志。

## 示例

```javascript
/*下例显示 Sheet1 中的第一个图表的第二个图例项的图例标示左边缘到图表区左边缘的距离。*/
function test() {
    let legend = Application.Worksheets.Item("Sheet1").ChartObjects(1).Chart.Legend
    console.log(legend.LegendEntries(2).LegendKey.Left)
}
```

```javascript
/*本示例将 Chart1 上第一个图例项的图例标示设置成有阴影。*/
function test() {
    Application.Charts.Item("Chart1").ChartObjects(1).Chart.Legend.LegendEntries(1).LegendKey.Shadow = true
}
```
