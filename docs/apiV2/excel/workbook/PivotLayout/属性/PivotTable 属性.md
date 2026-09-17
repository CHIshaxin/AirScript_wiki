---
title: "PivotTable 属性"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/PivotLayout/%E5%B1%9E%E6%80%A7/PivotTable%20%E5%B1%9E%E6%80%A7.html
---
# PivotTable (属性)

返回一个 **[PivotTable](./../../PivotTable/PivotTable 对象.html)** 对象，它代表与数据透视图相关联的数据透视表。

## 示例

```javascript
/*此示例将 Sheet1 上第一张图表对应的数据透视表的字段“id”的当前页设置为名为“3”的页。*/
function test() {
    let pvtTable = Worksheets.Item("Sheet1").ChartObjects(1).Chart.PivotLayout.PivotTable
    pvtTable.PivotFields("id").CurrentPage = "3"
}
```

```javascript
/*此示例确定与“Chart1”图表相关联的数据透视表，然后将字段“date”名为“2024/5/25”的页设置为数据透视表的当前页。*/
function test() {
    let pvtTable = Application.Charts.Item("Chart1").ChartObjects(1).Chart.PivotLayout.PivotTable
    pvtTable.PivotFields("date").CurrentPage = "2024/5/25"
}
```
