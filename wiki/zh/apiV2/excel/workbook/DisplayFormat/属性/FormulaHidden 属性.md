---
title: "FormulaHidden (属性)"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/DisplayFormat/%E5%B1%9E%E6%80%A7/FormulaHidden%20%E5%B1%9E%E6%80%A7.html
updated: 1789700916000
hash: BWoUiprs
---

# FormulaHidden (属性)

返回一个值，该值指示当前用户界面中显示的关联 **[Range](./../../Range/Range 对象.html)** 对象在工作表受保护时是否会隐藏自己的公式。只读。

## 说明

如果其公式在工作表受保护时处于隐藏状态，则返回 **True**；如果该区域中有些单元格的 **FormulaHidden** 等于 **True**，而有些单元格的 **FormulaHidden** 等于 **False**，则返回 **Null**。

## 返回值

Variant

## 示例

```javascript
/*此示例显示 D3 单元格在工作表受保护时是否隐藏公式。*/
function test() {
    console.log(Range("D3").DisplayFormat.FormulaHidden)
}
```

```javascript
/*此示例判断 E5:G5 单元格在工作表受保护时是否会隐藏公式，并通知用户。*/
function test() {
    if (Range("E5:G5").DisplayFormat.FormulaHidden) {
        console.log("工作表受保护时会隐藏公式")
    } else {
        console.log("工作表受保护时不会隐藏公式")
    }
}
```
