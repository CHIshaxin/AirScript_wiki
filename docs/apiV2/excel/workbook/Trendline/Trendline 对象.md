---
title: "Trendline 对象"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/Trendline/Trendline%20%E5%AF%B9%E8%B1%A1.html
---
# Trendline (对象)

代表图表上的趋势线。

## 说明

趋势线显示系列中数据的趋势或方向。 **[Trendline](./../Trendline/Trendline 对象.html)** 对象是 **[Trendlines](./../Trendlines/Trendlines 对象.html)** 集合的成员。 **[Trendlines](./../Trendlines/Trendlines 对象.html)** 集合包含某一个系列的所有 **[Trendline](./../Trendline/Trendline 对象.html)** 对象。

使用 **Trendlines**(*index*)（其中 *index*是趋势线索引号）可返回一个 **[Trendline](./../Trendline/Trendline 对象.html)** 对象。

索引号指出趋势线添加到系列中的顺序。`Trendlines(1)` 是第一个添加到系列中的趋势线，而 `Trendlines(Trendlines.Count)` 是最后一个。

## 示例

```javascript
/*下例更改工作表一上嵌入式图表一中第一个系列的趋势线类型。如果该系列没有趋势线，则本示例会失败。*/
function test() {
    Application.Worksheets.Item(1).ChartObjects(1).Chart.SeriesCollection(1).Trendlines(1).Type = xlMovingAvg
}
```

```javascript
/*本示例删除 Chart1 中第三个数据系列索引为2的趋势线。*/
function test() {
    let series = Application.Charts.Item("Chart1").ChartObjects(1).Chart.SeriesCollection(3)
    series.Trendlines(2).Delete()
}
```
