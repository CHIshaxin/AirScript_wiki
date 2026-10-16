---
title: "NumberFormatLinked 属性"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/DataLabel/%E5%B1%9E%E6%80%A7/NumberFormatLinked%20%E5%B1%9E%E6%80%A7.html
---
# NumberFormatLinked (属性)

如果指定数字格式指向单元格（以便当单元格的格式更改时数据标签的格式也作相应的改动），则为 **True**。**Boolean** 类型，可读写。

## 示例

```javascript
/*此示例显示工作表 Sheet1 中第一个图表的第二个数据序列的第一个数据标签的数字格式是否指向单元格。*/
function test() {
    let datalabel = Application.Sheets.Item("Sheet1").ChartObjects(1).Chart.SeriesCollection(2).DataLabels(1)
    console.log(datalabel.NumberFormatLinked)
}
```

```javascript
/*此示例将图表工作表 Chart1 中图表的第二个数据序列的第三个数据标签的数字格式设置为指向单元格。*/
function test() {
    let datalabel = Application.Charts.Item("Chart1").ChartObjects(1).Chart.SeriesCollection(2).DataLabels(3)
    datalabel.NumberFormatLinked  = true
}
```
