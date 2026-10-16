---
title: "Type 属性"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/Trendline/%E5%B1%9E%E6%80%A7/Type%20%E5%B1%9E%E6%80%A7.html
---
# Type (属性)

返回或设置一个 **XlTrendlineType**值，它代表趋势线的类型。

## 示例

```javascript
/*本示例更改第一个系列（第一张工作表的第一个嵌入图表中）的趋势线类型。如果该系列没有趋势线，则本示例无效。*/
function test() {
    Application.Worksheets.Item(1).ChartObjects(1).Chart.SeriesCollection(1).Trendlines(1).Type = xlMovingAvg
}
```

```javascript
/*下例更改 Chart1 中的第三个数据系列索引为2的趋势线的类型。*/
function test() {
    let trendline = Application.Charts.Item("Chart1").ChartObjects(1).Chart.SeriesCollection(3).Trendlines(2)
    trendline.Type = xlPolynomial
}
```
