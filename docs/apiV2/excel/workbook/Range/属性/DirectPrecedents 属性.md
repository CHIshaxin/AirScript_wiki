---
title: "DirectPrecedents 属性"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/Range/%E5%B1%9E%E6%80%A7/DirectPrecedents%20%E5%B1%9E%E6%80%A7.html
---
# DirectPrecedents (属性)

返回一个 **[Range](./../../Range/Range 对象.html)** 对象，该对象表示包含一个单元格的所有直接引用单元格的区域。如果有多个引用单元格，这可能有多个选择（多个 **[Range](./../../Range/Range 对象.html)** 对象）。 **[Range](./../../Range/Range 对象.html)** 对象，只读。

## 说明

**DirectPrecedents** 属性仅在活动工作表中有效，不能追踪远程引用。

## 示例

```javascript
/*本示例选定工作表 sheet1 中单元格 A1 的直接引用单元格。*/
function test() {
    Worksheets.Item("Sheet1").Activate()
    Range("A1").DirectPrecedents.Select()
}
```

```javascript
/*本示例在 A2 单元格设置公式 =AVERAGE(C2:C5)，并显示 A2 单元格的直接引用单元格地址*/
function test() {
    Range("A2").Formula = "=AVERAGE(C2:C5)"
    console.log(Range("A2").DirectPrecedents.Address())
}
```
