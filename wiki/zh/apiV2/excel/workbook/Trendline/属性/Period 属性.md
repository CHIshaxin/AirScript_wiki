---
title: "Period (属性)"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/Trendline/%E5%B1%9E%E6%80%A7/Period%20%E5%B1%9E%E6%80%A7.html
updated: 1789700916000
hash: BHIC1LKF
---

# Period (属性)

返回或设置移动平均趋势线的周期。可以是从 2 至 255 的值。可读/写 **Long** 类型。

## 示例

```javascript
/*本示例设置 Chart1 上移动平均趋势线的周期。本示例应在二维柱形图上运行，该柱形图中只有一个包含 10 个数据点的系列，并且有移动平均趋势线。*/
function test() {
    let trendline = Application.Charts.Item("Chart1").ChartObjects(1).Chart.SeriesCollection(1).Trendlines(1)
    if (trendline.Type == xlMovingAvg) {
        trendline.Period = 5
    }
}
```

```javascript
/*本示例设置第一张工作表上的第一个图表的第三个数据系列索引为2的移动平均趋势线的周期。*/
function test() {
    let trendline = Application.Worksheets.Item(1).ChartObjects(1).Chart.SeriesCollection(3).Trendlines(2)
    trendline.Period = 3
}
```
