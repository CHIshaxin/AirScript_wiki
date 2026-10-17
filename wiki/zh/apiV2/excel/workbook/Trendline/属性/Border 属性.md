---
title: "Border (属性)"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/Trendline/%E5%B1%9E%E6%80%A7/Border%20%E5%B1%9E%E6%80%A7.html
updated: 1789700916000
hash: CwLLsqXE
---

# Border (属性)

返回一个 **[Border](./../../Border/Border 对象.html)** 对象，它代表对象的边框。

## 示例

```javascript
/*下例更改 Chart1 中的第三个数据系列索引为1的趋势线线型*/
function test() {
    let trendline = Application.Charts.Item("Chart1").ChartObjects(1).Chart.SeriesCollection(3).Trendlines(1)
    trendline.Border.LineStyle = xlDash
}
```

```javascript
/*此示例显示第一张工作表上的第一个图表的第一个数据系列索引为2的趋势线边框粗细是否为细线。*/
function test() {
    let series1 = Application.Worksheets.Item(1).ChartObjects(1).Chart.SeriesCollection(1)
    console.log(series1.Trendlines(2).Border.Weight == xlHairline)
}
```
