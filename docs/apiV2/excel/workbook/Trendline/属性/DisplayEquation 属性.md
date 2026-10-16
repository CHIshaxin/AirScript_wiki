---
title: "DisplayEquation 属性"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/Trendline/%E5%B1%9E%E6%80%A7/DisplayEquation%20%E5%B1%9E%E6%80%A7.html
---
# DisplayEquation (属性)

如果显示图表中的趋势线公式，则该值为 **True**（其数据标签与 R-平方值相同）。将该属性设为 **True** 可自动显示数据标签。**Boolean**类型，可读写。

## 示例

```javascript
/*此示例显示 Chart1 上趋势线 1 的 R 平方值和公式。本示例应在具有第一个系列趋势线的 2D 柱形图上运行。*/
function test() {
    let trendline = Application.Charts.Item("Chart1").ChartObjects(1).Chart.SeriesCollection(1).Trendlines(1)
    trendline.DisplayRSquared = true
    trendline.DisplayEquation = true
}
```

```javascript
/*此示例显示第一张工作表上的第一个图表的第三个数据系列索引为1的趋势线公式。*/
function test() {
    let series3 = Application.Worksheets.Item(1).ChartObjects(1).Chart.SeriesCollection(3)
    series3.Trendlines(1).DisplayEquation = true
}
```
