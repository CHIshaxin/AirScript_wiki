---
title: "Select (方法)"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/DataLabels/%E6%96%B9%E6%B3%95/Select%20%E6%96%B9%E6%B3%95.html
updated: 1789700916000
hash: BXJ9xvsU
---

# Select (方法)

选择对象。

## 返回值

Variant

## 示例

```javascript
/*此示例选中图表工作表 Chart1 中图表的第一个数据序列的数据标签。*/
function test() {
    let datalabels = Application.Charts.Item("Chart1").ChartObjects(1).Chart.SeriesCollection(1).DataLabels()
    datalabels.Select()
}
```

```javascript
/*此示例选中工作表 Sheet1 中第一个图表的第三个数据序列的数据标签，并设置该数据标签的格式代码。*/
function test() {
    let datalabels = Application.Sheets.Item("Sheet1").ChartObjects(1).Chart.SeriesCollection(3).DataLabels()
    datalabels.Select()
    Selection.NumberFormat = "0.00"
}
```
