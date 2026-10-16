---
title: "DataTable (对象)"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/DataTable/DataTable%20%E5%AF%B9%E8%B1%A1.html
updated: 1789700916000
hash: ClxhDLOx
---

# DataTable (对象)

代表一张图表模拟运算表。

## 说明

使用 **DataTable** 属性可返回 **[DataTable](./../DataTable/DataTable 对象.html)** 对象。

## 示例

```javascript
/*此示例向嵌入式图表中添加带有外边框的模拟运算表。*/
function test() {
    let chart = Application.Worksheets.Item(1).ChartObjects(1).Chart
    chart.HasDataTable = true
    chart.DataTable.HasBorderOutline = true
}
```

```javascript
/*本示例将图表工作表 Chart1 中图表的模拟运算表中的边框和字体设置为蓝色。*/
function test() {
    let datatable = Application.Charts.Item("Chart1").ChartObjects(1).Chart.DataTable
    datatable.Border.ColorIndex = 5
    datatable.Font.ColorIndex = 5
}
```
