---
title: "ErrorString (属性)"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/PivotTable/%E5%B1%9E%E6%80%A7/ErrorString%20%E5%B1%9E%E6%80%A7.html
updated: 1789700916000
hash: IQ4-6Qj_
---

# ErrorString (属性)

返回或设置一个 **String** 值，它代表 **DisplayErrorString** 属性为 **True** 时如果单元格中有错误而显示的字符串。

## 说明

该属性的默认值是个空字符串 ("")。

## 示例

```javascript
/*对于指定数据透视表中有错误的单元格，本示例将在单元格中显示连字符。*/
function test() {
    Application.Worksheets.Item(1).PivotTables("Pivot1").ErrorString = "-"
    Application.Worksheets.Item(1).PivotTables("Pivot1").DisplayErrorString = true
}
```

```javascript
/*本示例指定数据透视表中有错误的单元格值显示为 error，并弹框展示该属性的值*/
function test() {
    ActiveSheet.PivotTables("Pivot1").DisplayErrorString = true
    ActiveSheet.PivotTables("Pivot1").ErrorString = "error"
    console.log(ActiveSheet.PivotTables("Pivot1").ErrorString)
}
```
