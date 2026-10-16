---
title: "FormulaHidden (属性)"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/Style/%E5%B1%9E%E6%80%A7/FormulaHidden%20%E5%B1%9E%E6%80%A7.html
updated: 1789700916000
hash: UqdqMZks
---

# FormulaHidden (属性)

返回或设置一个 **Boolean** 值，它指明在工作表处于保护状态时是否隐藏公式。

## 说明

请不要将此属性与 **Hidden** 属性混淆。如果工作簿受保护，而工作表不受保护，将不会隐藏公式。只有在工作表受保护时，才会隐藏公式。

## 示例

```javascript
/*本示例显示活动工作簿第一个样式是否隐藏公式（工作表处于保护状态时）。*/
function test() {
    console.log(ActiveWorkbook.Styles.Item(1).FormulaHidden)
}
```

```javascript
/*本示例在工作表处于保护状态时，隐藏C1单元格的公式。*/
function test() {
    Range("C1").Style.FormulaHidden = true
}
```
