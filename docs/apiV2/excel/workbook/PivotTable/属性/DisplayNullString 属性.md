---
title: "DisplayNullString 属性"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/PivotTable/%E5%B1%9E%E6%80%A7/DisplayNullString%20%E5%B1%9E%E6%80%A7.html
---
# DisplayNullString (属性)

如果数据透视表在包含空值的单元格中显示用户自定义的字符串，则该值为 **True**。默认值为 **True**。**Boolean** 类型，可读写。

## 说明

使用 **NullString** 属性可设置自定义空字符串。

## 示例

```javascript
/*本示例使数据透视表在包含空值的单元格中显示“NA”。*/
function test() {
    Application.Worksheets.Item(1).PivotTables("Pivot1").NullString = "NA"
    Application.Worksheets.Item(1).PivotTables("Pivot1").DisplayNullString = true
}
```

```javascript
/*本示例使数据透视表在包含空值的单元格中显示零 (0)。*/
function test() {
    Application.Worksheets.Item(1).PivotTables("Pivot1").DisplayNullString = false
}
```
