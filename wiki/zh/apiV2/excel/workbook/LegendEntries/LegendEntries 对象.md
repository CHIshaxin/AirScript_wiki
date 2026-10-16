---
title: "LegendEntries (对象)"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/LegendEntries/LegendEntries%20%E5%AF%B9%E8%B1%A1.html
updated: 1789700916000
hash: AmhtL0dh
---

# LegendEntries (对象)

指定的图表图例中所有 **[LegendEntry](./../LegendEntry/LegendEntry 对象.html)** 对象的集合。

## 说明

每个图例项都有两部分：一部分是该项的文本，它是与该图例项相关联的数据系列或趋势线的名称；另一部分是项标志，它在图表中以直观的方式将图例项以及与之相关联的数据系列或趋势线链接起来。项标志及其相关系列或趋势线的格式属性包含在 **[LegendKey](./../LegendKey/LegendKey 对象.html)** 对象中。

## 示例

```javascript
/*下例在嵌入式图表一中的图例项集合中循环，并更改这些图例项的字体颜色。*/
function test() {
    let lgd = Application.Worksheets.Item("Sheet1").ChartObjects(1).Chart.Legend
    for (let i = 1; i <= lgd.LegendEntries().Count; i++) {
        lgd.LegendEntries(i).Font.ColorIndex = 5
    }
}
```

使用 **LegendEntries**( *index*)（其中 *index* 是图例项索引号）可返回一个 **[LegendEntry](./../LegendEntry/LegendEntry 对象.html)** 对象。不能按名称返回图例项。

图例项的编号代表图例项在图例中的位置。`LegendEntries(1)` 位于图例的顶部；`LegendEntries(LegendEntries.Count)` 位于图例的底部。

```javascript
/*下例将嵌入的第一个图表中图例顶部的图例项（这通常是第一个数据系列的图例）的文字字体设置为斜体。*/
function test() {
    Application.Worksheets.Item("Sheet1").ChartObjects(1).Chart.Legend.LegendEntries(1).Font.Italic = true
}
```
