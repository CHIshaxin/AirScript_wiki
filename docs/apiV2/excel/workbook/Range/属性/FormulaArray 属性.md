---
title: "FormulaArray 属性"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/Range/%E5%B1%9E%E6%80%A7/FormulaArray%20%E5%B1%9E%E6%80%A7.html
---
# FormulaArray (属性)

返回或设置区域的数组公式。返回（或可设置为）单个公式或数组。如果指定区域不包含数组公式，则该属性返回 **null**。**Variant** 类型，可读写。

## 说明

如果用本属性输入数组公式，那么该公式必须使用 R1C1 引用样式，而不能使用 A1 引用样式（请参阅第二个示例）。

**FormulaArray** 属性的字符限制为 255。

## 示例

```javascript
/*本示例在 Sheet1 上的单元格区域 A1:C5 中输入数组常量 3。*/
function test() {
    Worksheets.Item("Sheet1").Range("A1:C5").FormulaArray = "=3"
}
```

```javascript
/*本示例在 sheet1 上的单元格区域 E1:E3 中输入数组公式“=SUM(R1C1:R3C3)”。*/
function test() {
    Worksheets.Item("Sheet1").Range("E1:E3").FormulaArray = "=Sum(R1C1:R3C3)"
}
```
