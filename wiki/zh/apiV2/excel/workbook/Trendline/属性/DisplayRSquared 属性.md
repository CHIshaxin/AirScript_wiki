---
title: "DisplayRSquared (属性)"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/Trendline/%E5%B1%9E%E6%80%A7/DisplayRSquared%20%E5%B1%9E%E6%80%A7.html
updated: 1789700916000
hash: I-kN1gHO
---

# DisplayRSquared (属性)

如果显示图表中趋势线的 R-平方值，则该值为 **True**（其数据标签与公式的相同）。将该属性设为 **True** 可自动显示数据标签。**Boolean**类型，可读写。

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
/*此示例显示第一张工作表上的第一个图表的第二个数据系列索引为2的趋势线的 R 平方值。*/
function test() {
    let series2 = Application.Worksheets.Item(1).ChartObjects(1).Chart.SeriesCollection(2)
    series2.Trendlines(2).DisplayRSquared = true
}
```
