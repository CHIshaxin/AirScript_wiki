---
title: "Select 方法"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/DataTable/%E6%96%B9%E6%B3%95/Select%20%E6%96%B9%E6%B3%95.html
---
# Select (方法)

选择对象。

## 示例

```javascript
/*本示例选中图表工作表 Chart1 中图表的模拟运算表。*/
function test() {
    let datatable = Application.Charts.Item("Chart1").ChartObjects(1).Chart.DataTable
    datatable.Select()
}
```

```javascript
/*本示例选中工作表 Sheet1 中第一张图表的模拟运算表，并设置边框颜色。*/
function test() {
    let datatable = Application.Sheets.Item("Sheet1").ChartObjects(1).Chart.DataTable
    datatable.Select()
    Selection.Border.ColorIndex = 4
}
```
