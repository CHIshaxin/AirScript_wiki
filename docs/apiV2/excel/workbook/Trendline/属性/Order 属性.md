---
title: "Order 属性"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/Trendline/%E5%B1%9E%E6%80%A7/Order%20%E5%B1%9E%E6%80%A7.html
---
# Order (属性)

返回或设置一个 **Long** 值，它代表趋势线类型为 **xlPolynomial** 时趋势线的顺序号（大于 1 的整数）。

## 示例

```javascript
/*此示例显示第一张工作表上的第一个图表的第二个数据系列索引为2的趋势线顺序。*/
function test() {
    let series = Application.Worksheets.Item(1).ChartObjects(1).Chart.SeriesCollection(2)
    console.log(series.Trendlines(2).Order)
}
```

```javascript
/*本示例设置 Chart1 中第三个数据系列索引为1的趋势线顺序为3。*/
function test() {
    let series = Application.Charts.Item("Chart1").ChartObjects(1).Chart.SeriesCollection(3)
    series.Trendlines(1).Order = 3
}
```
