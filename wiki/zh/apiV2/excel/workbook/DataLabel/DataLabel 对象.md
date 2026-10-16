---
title: "DataLabel (对象)"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/DataLabel/DataLabel%20%E5%AF%B9%E8%B1%A1.html
updated: 1789700916000
hash: BCxrWu7z
---

# DataLabel (对象)

代表图表数据点或趋势线上的数据标签。

## 说明

在数据系列上， **[DataLabel](./../DataLabel/DataLabel 对象.html)** 对象是 **[DataLabels](./../DataLabels/DataLabels 对象.html)** 集合的成员。 **[DataLabels](./../DataLabels/DataLabels 对象.html)** 集合包含每个数据点的 **[DataLabel](./../DataLabel/DataLabel 对象.html)** 对象。对于没有可定义数据点的数据系列（如面积图数据系列）， **[DataLabels](./../DataLabels/DataLabels 对象.html)** 集合包含单个 **[DataLabel](./../DataLabel/DataLabel 对象.html)** 对象。

使用 **DataLabels**( *index*)（其中 *index* 为数据标签的索引号）可返回单个 **[DataLabel](./../DataLabel/DataLabel 对象.html)** 对象。

## 示例

```javascript
/*本示例在工作表 Sheet1 上第一个图表上，设置第一个数据系列中的第五个数据标签的数字格式。*/
function test() {
    let datalabel = Application.Worksheets.Item("Sheet1").ChartObjects(1).Chart.SeriesCollection(1).DataLabels(5)
    datalabel.NumberFormat = "0.000"
}
```

使用 **DataLabel** 属性可返回单个数据点的 **[DataLabel](./../DataLabel/DataLabel 对象.html)** 对象。

```javascript
/*此示例打开名为“Chart1”的图表工作表上第一个数据系列中第二个数据点的数据标签，并将数据标签文本设置为“Saturday”。*/
function test() {
    let point = Application.Charts.Item("Chart1").ChartObjects(1).Chart.SeriesCollection(1).Points(2)
    point.HasDataLabel = true
    point.DataLabel.Text = "Saturday"
}
```

在趋势线上，**DataLabel**返回与趋势线一起显示的文本。这些文本可能是公式、R 平方值或两者均有（如果两者都出现）。

```javascript
/*本示例将趋势线文本设置为仅显示公式，并将数据标签的名称置于工作表“Sheet1”上的单元格 A1 中。*/
function test() {
    let trendline = Application.Charts.Item("Chart1").ChartObjects(1).Chart.SeriesCollection(1).Trendlines(1)
    trendline.DisplayRSquared = false
    trendline.DisplayEquation = true
    Worksheets.Item("Sheet1").Range("A1").Value2 = trendline.DataLabel.Name
}
```
