---
title: "NullString 属性"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/PivotTable/%E5%B1%9E%E6%80%A7/NullString%20%E5%B1%9E%E6%80%A7.html
---
# NullString (属性)

返回或设置当 **DisplayNullString** 为 **True** 时，在包含 null 值的单元格中显示的字符串。默认值为空字符串 ("")。**String** 类型，可读写。

## 示例

```javascript
/*本示例使数据透视表在包含空值的单元格中显示“NA”。*/
function test() {
    Application.Worksheets.Item(1).PivotTables("Pivot1").NullString = "NA"
    Application.Worksheets.Item(1).PivotTables("Pivot1").DisplayNullString = true
}
```

```javascript
/*本示例使数据透视表在包含空值的单元格中显示 NULL，并弹窗显示此属性*/
function test() {
    ActiveSheet.PivotTables("Pivot1").DisplayNullString = true
    ActiveSheet.PivotTables("Pivot1").NullString = "NULL"
    console.log(ActiveSheet.PivotTables("Pivot1").NullString)
}
```
