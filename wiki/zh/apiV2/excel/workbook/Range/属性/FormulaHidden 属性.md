---
title: "FormulaHidden (属性)"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/Range/%E5%B1%9E%E6%80%A7/FormulaHidden%20%E5%B1%9E%E6%80%A7.html
updated: 1789700916000
hash: wK0UNLRG
---

# FormulaHidden (属性)

返回或设置一个 **Variant** 值，它指明在工作表处于保护状态时是否隐藏公式。

## 说明

如果在工作表处于保护状态时要隐藏公式，此属性将返回 **True**；如果指定区域中有些单元格的 **FormulaHidden** 为 **True**，而有些单元格的 **FormulaHidden** 为 **False**，则返回 **null**。

请不要将此属性与 **Hidden** 属性混淆。如果工作簿受保护，而工作表不受保护，将不会隐藏公式。只有在工作表受保护时，才会隐藏公式。

## 示例

```javascript
/*此示例在工作表处于保护状态时，隐藏 Sheet1 中单元格 A1 和 B1 的公式。*/
function test() {
    Worksheets.Item("Sheet1").Range("A1:B1").FormulaHidden = true
}
```

```javascript
/*本示例工作表处于保护状态时，隐藏 Sheet1 中单元格 A1 和 B1 的公式，显示 A1:C1 单元格公式的隐藏状态*/
function test() {
    Worksheets.Item("Sheet1").Range("A1:B1").FormulaHidden = true
    console.log(Range("A1").FormulaHidden)
}
```
