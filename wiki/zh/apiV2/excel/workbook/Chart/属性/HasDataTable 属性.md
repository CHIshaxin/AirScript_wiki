---
title: "HasDataTable (属性)"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/Chart/%E5%B1%9E%E6%80%A7/HasDataTable%20%E5%B1%9E%E6%80%A7.html
updated: 1789700916000
hash: DA1gfqb1
---

# HasDataTable (属性)

如果图表有模拟运算表，则该属性值为 **True**。**Boolean** 类型，可读写。

## 示例

```javascript
/*本示例使第一张工作表第一个嵌入式图表模拟运算表显示时带有外边框，但无单元格边框。*/
function test() {
    let chart = Application.Worksheets.Item(1).ChartObjects(1).Chart
    chart.HasDataTable = true
    let table = chart.DataTable
    table.HasBorderHorizontal = false
    table.HasBorderVertical = false
    table.HasBorderOutline = true
}
```

```javascript
/*本示例判断图表工作表是否拥有模拟运算表，并通知用户。*/
function test() {
    let chart = Application.Charts.Item("Chart1")
    if (chart.HasDataTable) {
        console.log("该图表工作表有模拟运算表")
    } else {
        console.log("该图表工作表没有模拟运算表")
    }
}
```
