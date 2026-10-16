---
title: "Format 属性"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/DataLabel/%E5%B1%9E%E6%80%A7/Format%20%E5%B1%9E%E6%80%A7.html
---
# Format (属性)

返回 **[ChartFormat](./../../ChartFormat/ChartFormat 对象.html)** 对象。只读。

## 示例

```javascript
/*此示例将工作表 Sheet1 中第一个图表的第一个数据序列的第一个数据标签的边框颜色设置为红色。*/
function test() {
    let datalabel = Application.Sheets.Item("Sheet1").ChartObjects(1).Chart.SeriesCollection(1).DataLabels(1)
    datalabel.Format.Line.ForeColor.RGB = RGB(255, 0, 0)
}
```

```javascript
/*此示例将图表工作表 Chart1 中图表的第二个数据序列的第一个数据标签的前景色设置为蓝色。*/
function test() {
    let datalabel = Application.Charts.Item("Chart1").ChartObjects(1).Chart.SeriesCollection(2).DataLabels(1)
    datalabel.Format.Fill.ForeColor.RGB = RGB(0, 255, 255)
}
```
