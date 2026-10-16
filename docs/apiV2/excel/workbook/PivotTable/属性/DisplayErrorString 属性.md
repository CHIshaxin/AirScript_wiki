---
title: "DisplayErrorString 属性"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/PivotTable/%E5%B1%9E%E6%80%A7/DisplayErrorString%20%E5%B1%9E%E6%80%A7.html
---
# DisplayErrorString (属性)

如果数据透视表在有错误的单元格中显示用户自定义的错误字符串，则该值为 **True**。默认值为 **False**。**Boolean** 类型，可读写。

## 说明

使用 **ErrorString**属性可设置用户自定义的错误字符串。

当透视处理计算字段时，该属性对于消除被零除错误尤其有用。

## 示例

```javascript
/*本示例使数据透视表在有错误的单元格中显示连字符。*/
function test() {
    Worksheets.Item(1).PivotTables("Pivot1").ErrorString = "-"
    Worksheets.Item(1).PivotTables("Pivot1").DisplayErrorString = true
}
```

```javascript
/*本示例显示工作表 Sheet1 中第一张数据透视表的 DisplayErrorString 属性的数据类型，并将其设置为 false*/
function test() {
    console.log(typeof Worksheets.Item("Sheet1").PivotTables(1).DisplayErrorString)
    Worksheets.Item("Sheet1").PivotTables(1).DisplayErrorString = false
}
```
