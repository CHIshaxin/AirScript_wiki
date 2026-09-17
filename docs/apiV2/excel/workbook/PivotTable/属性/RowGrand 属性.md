---
title: "RowGrand 属性"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/PivotTable/%E5%B1%9E%E6%80%A7/RowGrand%20%E5%B1%9E%E6%80%A7.html
---
# RowGrand (属性)

如果数据透视表显示行的总数，则该属性值为 **True**。**Boolean** 类型，可读写。

## 示例

```javascript
/*本示例设置数据透视表以显示行的总计。*/
function test() {
    let pvtTable = Application.Worksheets.Item("Sheet1").Range("F1").PivotTable
    pvtTable.RowGrand = true
}
```

```javascript
/*本示例显示第一张数据透视表 RowGrand 属性的数据类型，并设置为 false。*/
function test() {
    console.log(typeof ActiveSheet.PivotTables(1).RowGrand)
    ActiveSheet.PivotTables(1).RowGrand = false
}
```
