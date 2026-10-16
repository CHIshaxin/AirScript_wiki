---
title: "Formula 属性"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/DataLabel/%E5%B1%9E%E6%80%A7/Formula%20%E5%B1%9E%E6%80%A7.html
---
# Formula (属性)

获取或设置一个 **String** 值，该值以英语使用 A1 样式表示法来表示对象的公式。可读/写。

## 示例

```javascript
/*此示例将图表工作表 Chart1 中图表的第二个数据序列的第一个数据标签设置为工作表 Sheet1 中单元格 E1 的内容。*/
function test() {
    let datalabel = Application.Charts.Item("Chart1").ChartObjects(1).Chart.SeriesCollection(2).DataLabels(1)
    datalabel.Formula = "=Sheet1!$E$1"
}
```

```javascript
/*此示例将工作表 Sheet1 中第一个图表的第一个数据序列的第二个数据标签设置为该工作表单元格 A1 的内容。*/
function test() {
    let datalabel = Application.Sheets.Item("Sheet1").ChartObjects(1).Chart.SeriesCollection(1).DataLabels(2)
    datalabel.Formula = "=Sheet1!$A$1"
}
```
