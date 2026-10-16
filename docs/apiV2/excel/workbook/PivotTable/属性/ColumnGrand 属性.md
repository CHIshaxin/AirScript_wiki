---
title: "ColumnGrand 属性"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/PivotTable/%E5%B1%9E%E6%80%A7/ColumnGrand%20%E5%B1%9E%E6%80%A7.html
---
# ColumnGrand (属性)

如果数据透视表显示列总计，则该属性值为 **True**。**Boolean** 类型，可读写。

## 示例

```javascript
/*本示例设置数据透视表显示列总计。*/
function test() {
    let pvtTable = Worksheets.Item("Sheet1").Range("A3").PivotTable
    pvtTable.ColumnGrand = true
}
```

```javascript
/*本示例显示活动工作表中第一张数据透视表 ColumnGrand 属性的数据类型，并将其设置为 false。*/
function test() {
    console.log(`第一张数据透视表中ColumnGrand属性的数据类型是：${typeof ActiveSheet.PivotTables(1).ColumnGrand}`)
    ActiveSheet.PivotTables(1).ColumnGrand = false
}
```
