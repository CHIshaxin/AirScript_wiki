---
title: "ClearFormats (方法)"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/Trendline/%E6%96%B9%E6%B3%95/ClearFormats%20%E6%96%B9%E6%B3%95.html
updated: 1789700916000
hash: J-S7ehKb
---

# ClearFormats (方法)

清除对象的格式设置。

## 返回值

Variant

## 示例

```javascript
/*下例清除 Chart1 中的第三个数据系列索引为1的趋势线的格式*/
function test() {
    let trendline = Application.Charts.Item("Chart1").ChartObjects(1).Chart.SeriesCollection(3).Trendlines(1)
    trendline.ClearFormats()
}
```

```javascript
/*此示例清除第一张工作表上的第一个图表的第二个数据系列索引为2的趋势线格式。*/
function test() {
    let series = Application.Worksheets.Item(1).ChartObjects(1).Chart.SeriesCollection(2)
    series.Trendlines(2).ClearFormats()
}
```
