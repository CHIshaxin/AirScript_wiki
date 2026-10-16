---
title: "DataLabel 属性"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/Trendline/%E5%B1%9E%E6%80%A7/DataLabel%20%E5%B1%9E%E6%80%A7.html
---
# DataLabel (属性)

返回一个 **[DataLabel](./../../DataLabel/DataLabel 对象.html)** 对象，它代表与趋势线相关的数据标签。只读。

## 说明

若要打开趋势线的数据标签，需要将 **DisplayEquation** 属性和 **DisplayRSquared** 属性设置为 **True**。

## 示例

```javascript
/*下例为 Chart1 中的第一个数据系列索引为1的趋势线启用数据标签，并显示数据标签文本方向是否为水平排列*/
function test() {
    let trendline = Application.Charts.Item("Chart1").ChartObjects(1).Chart.SeriesCollection(1).Trendlines(1)
    trendline.DisplayRSquared = true
    trendline.DisplayEquation = true
    console.log(trendline.DataLabel.Orientation == xlHorizontal)
}
```

```javascript
/*此示例为第一张工作表上的第一个图表的第二个数据系列索引为2的趋势线启用数据标签，并选中数据标签。*/
function test() {
    let series = Application.Worksheets.Item(1).ChartObjects(1).Chart.SeriesCollection(2)
    series.Trendlines(2).DisplayRSquared = true
    series.Trendlines(2).DisplayEquation = true
    series.Trendlines(2).DataLabel.Select()
}
```
