---
title: "FormulaHidden (属性)"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/CellFormat/%E5%B1%9E%E6%80%A7/FormulaHidden%20%E5%B1%9E%E6%80%A7.html
updated: 1789700916000
hash: BiGi2tNv
---

# FormulaHidden (属性)

返回或设置一个 **Variant** 值，它指明在工作表处于保护状态时是否隐藏公式。

## 说明

如果在工作表处于保护状态时要隐藏公式，此属性将返回 **True**；如果指定区域中有些单元格的 **FormulaHidden** 为 **True**，而有些单元格的 **FormulaHidden** 为 **False**，则返回 **Null**。

请不要将此属性与 **Hidden** 属性混淆。如果工作簿受保护，而工作表不受保护，将不会隐藏公式。只有在工作表受保护时，才会隐藏公式。

## 示例

```javascript
/*此示例将查找条件中工作表（保护状态时）设置为隐藏公式。*/
function test() {
    Application.FindFormat.FormulaHidden = true
}
```

```javascript
/*此示例显示替换条件中工作表处于保护状态时是否隐藏公式。*/
function test() {
    console.log(Application.ReplaceFormat.FormulaHidden)
}
```
