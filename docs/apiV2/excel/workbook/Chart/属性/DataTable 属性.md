---
title: "DataTable 属性"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/Chart/%E5%B1%9E%E6%80%A7/DataTable%20%E5%B1%9E%E6%80%A7.html
---
# DataTable (属性)

返回一个 **[DataTable](./../../DataTable/DataTable 对象.html)** 对象，该对象表示图表的模拟运算表。只读。

## 示例

```javascript
/*本示例向活动工作表第一张嵌入图表添加带有外框的模拟运算表。*/
function test() {
    let chart = Application.ActiveSheet.ChartObjects(1).Chart
    chart.HasDataTable = true
    chart.DataTable.HasBorderOutline = true
}
```

```javascript
/*本示例使第一张工作表第一个内嵌图表模拟运算表显示时带有外边框，但无单元格边框。*/
function test() {
    let chart = Application.Sheets.Item(1).ChartObjects(1).Chart
    chart.HasDataTable = true
    let datatable = chart.DataTable
    datatable.HasBorderHorizontal = false
    datatable.HasBorderVertical = false
    datatable.HasBorderOutline = true
}
```
