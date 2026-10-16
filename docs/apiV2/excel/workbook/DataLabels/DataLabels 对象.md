---
title: "DataLabels 对象"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/DataLabels/DataLabels%20%E5%AF%B9%E8%B1%A1.html
---
# DataLabels (对象)

由数据系列中所有 **[DataLabel](./../DataLabel/DataLabel 对象.html)** 对象组成的集合。

## 说明

每个 **[DataLabel](./../DataLabel/DataLabel 对象.html)** 对象代表一个数据点或趋势线的数据标签。对于没有可定义数据点的数据系列（例如面积图数据系列）， **[DataLabels](./../DataLabels/DataLabels 对象.html)** 集合包含单个数据标签。

使用 **DataLabels** 方法可返回单个 **[DataLabels](./../DataLabels/DataLabels 对象.html)** 集合。

## 示例

```javascript
/*此示例设置图表工作表 Chart1 中图表的第一个数据系列中数据标签的数字格式。*/
function test() {
    let series = Application.Charts.Item("Chart1").ChartObjects(1).Chart.SeriesCollection(1)
    series.HasDataLabels = true
    series.DataLabels().NumberFormat = "#.##0"
}
```

使用 **DataLabels**(*index*)（其中 index 为数据标签的索引号）可返回单个 **[DataLabel](./../DataLabel/DataLabel 对象.html)** 对象。

```javascript
/*此示例在第一张工作表上嵌入的第一个图表上，设置第一个数据系列中的第五个数据标签的数字格式。*/
function test() {
    let datalabel = Application.Worksheets.Item(1).ChartObjects(1).Chart.SeriesCollection(1).DataLabels(5)
    datalabel.NumberFormat = "0.000"
}
```
