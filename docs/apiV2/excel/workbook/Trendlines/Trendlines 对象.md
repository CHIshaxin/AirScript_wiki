---
title: "Trendlines 对象"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/Trendlines/Trendlines%20%E5%AF%B9%E8%B1%A1.html
---
# Trendlines (对象)

指定的数据系列中所有 **[Trendline](./../Trendline/Trendline 对象.html)** 对象的集合。

## 说明

每一个 **[Trendline](./../Trendline/Trendline 对象.html)** 对象都代表图表中的趋势线。趋势线显示系列中数据的趋势或方向。

## 示例

```javascript
/*以下示例显示 Chart1 上第一系列的趋势线数。*/
function test() {
    console.log(Charts.Item("Chart1").ChartObjects(1).Chart.SeriesCollection(1).Trendlines().Count)
}
```

```javascript
/*下例给 Sheet1 中嵌入式图表一中的第一个系列添加线性趋势线。*/
function test() {
    Worksheets("Sheet1").ChartObjects(1).Chart.SeriesCollection(1).Trendlines().Add(xlLinear, null, null, null, null, null, null, null, "Linear Trend")
}
```

```javascript
/*下例更改工作表一上嵌入式图表一中第一个系列的趋势线类型。如果该系列没有趋势线，本示例将失败。*/
function test() {
    Worksheets(1).ChartObjects(1).Chart.SeriesCollection(1).Trendlines(1).Type = xlMovingAvg
}
```
