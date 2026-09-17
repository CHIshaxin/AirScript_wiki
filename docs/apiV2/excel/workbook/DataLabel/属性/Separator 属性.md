---
title: "Separator 属性"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/DataLabel/%E5%B1%9E%E6%80%A7/Separator%20%E5%B1%9E%E6%80%A7.html
---
# Separator (属性)

设置或返回一个 **Variant** 值，它代表用于图表中数据标签的分隔符。可读/写。

## 说明

如果使用字符串，则用一个字符串作为分隔符。如果使用 **xlDataLabelSeparatorDefault** (= 1)，则用默认数据标签分隔符，即逗号或换行符，具体取决于数据标签。

当返回值“1”时，表明用户未更改默认的分隔符逗号“,”。您也可通过传递值“1”将分隔符更改回默认值。

以程序方式访问数据标签之前必须先激活图表，否则将出现运行错误。

## 示例

```javascript
/*此示例将工作表 Sheet1 中第一个图表上的第一个系列的第一个数据标签分隔符设置为分号。此示例假定工作表上存在图表。*/
function test() {
    let datalabel = Application.Worksheets.Item("Sheet1").ChartObjects(1).Chart.SeriesCollection(1).DataLabels(1)
    datalabel.Separator = ";"
}
```

```javascript
/*此示例将图表工作表 Chart1 中图表的第一个数据序列的第二个数据标签的分隔符设置为默认值。*/
function test() {
    let datalabel = Application.Charts.Item("Chart1").ChartObjects(1).Chart.SeriesCollection(1).DataLabels(2)
    datalabel.Separator = xlDataLabelSeparatorDefault
}
```
