---
title: "PreserveFormatting 属性"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/PivotTable/%E5%B1%9E%E6%80%A7/PreserveFormatting%20%E5%B1%9E%E6%80%A7.html
---
# PreserveFormatting (属性)

当使用数据透视、筛选或更改页字段项等操作刷新或重新计算报告时，如果格式处于被保护状态，则为 **True**。对于查询表，如果数据前五行的任何常规格式设置将应用于查询表数据的新行，则此属性为 **True**。对未使用的单元格不进行设置。如果将查询表应用的最近一次自动套用格式应用于新数据行，则此属性为 **False**。默认值是 **True**。

## 说明

对于数据库查询表，默认的格式设置为 **xlSimple**。

刷新查询表时，将对查询表应用新的自动套用格式样式。只要 **PreserveFormatting** 的值为 **False**，则 AutoFormat（自动套用格式）就会被设置为 **None**。这样，任何在 **PreserveFormatting** 被设置为 **False**或在查询表刷新之前设置的自动套用格式都不会起作用，且相应产生的查询表也不会被应用任何格式。

## 示例

```javascript
/*此示例保留第一张工作表上的第一个数据透视表的格式。*/
function test() {
    Application.Worksheets.Item(1).PivotTables("Pivot1").PreserveFormatting = true
}
```

```javascript
/*本示例显示数据透视表 PreserveFormatting 属性的数据类型，并设置为 false。*/
function test() {
    let pvtTable = ActiveSheet.PivotTables(1)
    console.log(typeof pvtTable.PreserveFormatting)
    pvtTable.PreserveFormatting = false
    pvtTable.RefreshTable()
}
```
