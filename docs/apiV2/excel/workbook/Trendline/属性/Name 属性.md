---
title: "Name 属性"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/Trendline/%E5%B1%9E%E6%80%A7/Name%20%E5%B1%9E%E6%80%A7.html
---
# Name (属性)

返回或设置一个 **String** 值，它代表对象的名称。

## 示例

```javascript
/*本示例设置 Chart1 中第一个数据系列索引为1的趋势线名称。*/
function test() {
    let series1 = Application.Charts.Item("Chart1").ChartObjects(1).Chart.SeriesCollection().Item(1)
    series1.Trendlines(1).Name = "趋势线(系列1)"
}
```

```javascript
/*此示例显示第一张工作表上的第一个图表的第二个数据系列索引为1的趋势线名称。*/
function test() {
    let series2 = Application.Worksheets.Item(1).ChartObjects(1).Chart.SeriesCollection(2)
    console.log(series2.Trendlines(1).Name)
}
```
