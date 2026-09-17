---
title: "HasBorderHorizontal 属性"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/DataTable/%E5%B1%9E%E6%80%A7/HasBorderHorizontal%20%E5%B1%9E%E6%80%A7.html
---
# HasBorderHorizontal (属性)

如果图表模拟运算表具有水平单元格边框，则该属性值为 **True**。**Boolean** 类型，可读写。

## 示例

```javascript
/*本示例使嵌入图表模拟运算表显示时带有外边框，但无单元格边框。*/
function test() {
    let chart = Application.Worksheets.Item(1).ChartObjects(1).Chart
    chart.HasDataTable = true
    let datatable = chart.DataTable
    datatable.HasBorderHorizontal = false
    datatable.HasBorderVertical = false
    datatable.HasBorderOutline = true
}
```

```javascript
/*本示例显示图表工作表 Chart1 中图表的模拟运算表是否具有水平单元格边框。*/
function test() {
    let datatable = Application.Charts.Item("Chart1").ChartObjects(1).Chart.DataTable
    console.log(datatable.HasBorderHorizontal)
}
```
