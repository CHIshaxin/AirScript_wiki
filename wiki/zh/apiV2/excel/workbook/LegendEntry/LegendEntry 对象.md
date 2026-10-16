---
title: "LegendEntry (对象)"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/LegendEntry/LegendEntry%20%E5%AF%B9%E8%B1%A1.html
updated: 1789700916000
hash: CfAq-ZsG
---

# LegendEntry (对象)

代表图表图例中的图例项。

## 说明

**[LegendEntry](./../LegendEntry/LegendEntry 对象.html)** 对象是 **[LegendEntries](./../LegendEntries/LegendEntries 对象.html)** 集合的成员。 **[LegendEntries](./../LegendEntries/LegendEntries 对象.html)** 集合包含图例中所有的 **[LegendEntry](./../LegendEntry/LegendEntry 对象.html)** 对象。

每个图例项都有两部分：一部分是该项的文本，它是与该图例项相关联的数据系列或趋势线的名称；另一部分是项标志，它在图表中以直观的方式将图例项以及与之相关联的数据系列或趋势线链接起来。项标志及其相关数据系列或趋势线的格式设置属性包含在 **[LegendKey](./../LegendKey/LegendKey 对象.html)** 对象中。

不能修改图例项的文字。 **[LegendEntry](./../LegendEntry/LegendEntry 对象.html)** 对象支持字体格式，且可被删除。图例项不支持图案格式。图例项的位置和尺寸是固定的。

没有返回相应于某图例项的数据系列或趋势线的直接方法。

在删除了图例项之后，唯一的恢复方法是：通过将图表的 **HasLegend** 属性设置为 **False**，然后再设回 **True**，从而删除并重新创建包含这些图例项的图例。

使用 **LegendEntries**( *index*)（其中 *index*是图例项索引号）可返回一个 **[LegendEntry](./../LegendEntry/LegendEntry 对象.html)** 对象。不能按名称返回图例项。

图例项的编号代表图例项在图例中的位置。`LegendEntries(1)` 位于图例的顶部； `LegendEntries(LegendEntries.Count)` 位于图例的底部。

## 示例

```javascript
/*下例更改工作表“Sheet1”上嵌入的第一个图表中图例顶部的图例项（这通常是第一个数据系列的图例）的文字字体。*/
function test() {
    Application.Worksheets.Item("Sheet1").ChartObjects(1).Chart.Legend.LegendEntries(1).Font.Italic = true
}
```

```javascript
/*下例显示 Chart1 上第二个图例项的高度。*/
function test() {
    let legend = Application.Charts.Item("Chart1").ChartObjects(1).Chart.Legend
    console.log(legend.LegendEntries(2).Height)
}
```
