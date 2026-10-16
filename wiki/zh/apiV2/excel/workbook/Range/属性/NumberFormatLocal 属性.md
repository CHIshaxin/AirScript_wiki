---
title: "NumberFormatLocal (属性)"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/Range/%E5%B1%9E%E6%80%A7/NumberFormatLocal%20%E5%B1%9E%E6%80%A7.html
updated: 1789700916000
hash: CJEvxSZW
---

# NumberFormatLocal (属性)

以采用用户语言字符串的形式返回或设置一个 **Variant** 值，它代表对象的格式代码。

## 说明

**Format** 函数使用的格式代码字符串与 **NumberFormat** 和 **NumberFormatLocal** 属性使用的格式代码字符串不同。

## 示例

```javascript
/*本示例以用户语言显示 Sheet1 上的单元格 A1 中的数字格式。*/
function test() {
    let numberFormula = Worksheets.Item("Sheet1").Range("A1").NumberFormatLocal
    console.log("The number format for cell A1 is " + numberFormula)
}
```

```javascript
/*本示例将 A1:C4 单元格区域的数字格式（用户语言）设置为 yyyy/m/d，并显示，然后将 E1 单元格的数字格式（用户语言）设置为 # ??/??*/
function test() {
    Range("A1:C4").NumberFormatLocal = "yyyy/m/d"
    console.log(Range("C4").NumberFormatLocal)
    Range("E1").NumberFormatLocal = "# ??/??"
}
```
