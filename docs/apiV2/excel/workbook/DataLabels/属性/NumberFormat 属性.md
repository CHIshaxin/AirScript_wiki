---
title: "NumberFormat 属性"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/DataLabels/%E5%B1%9E%E6%80%A7/NumberFormat%20%E5%B1%9E%E6%80%A7.html
---
# NumberFormat (属性)

返回或设置一个 **String** 值，它代表对象的格式代码。

## 说明

格式代码与**设置单元格格式**对话框中的**格式代码**选项是同一个字符串。**Format** 函数使用的格式代码字符串与 **NumberFormat** 和 **NumberFormatLocal** 属性使用的格式代码字符串不同。

## 示例

```javascript
/*此示例设置图表工作表 Chart1 中图表的第二个数据序列的数据标签的格式代码。*/
function test() {
    let datalabels = Application.Charts.Item("Chart1").ChartObjects(1).Chart.SeriesCollection(2).DataLabels()
    datalabels.NumberFormat = "0.00"
}
```

```javascript
/*此示例设置工作表 Sheet1 中第一个图表的第一个数据序列的数据标签的格式代码。*/
function test() {
    let datalabels = Application.Sheets.Item("Sheet1").ChartObjects(1).Chart.SeriesCollection(1).DataLabels()
    datalabels.NumberFormat = "#.##0"
}
```
