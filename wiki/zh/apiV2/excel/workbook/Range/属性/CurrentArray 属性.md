---
title: "CurrentArray (属性)"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/Range/%E5%B1%9E%E6%80%A7/CurrentArray%20%E5%B1%9E%E6%80%A7.html
updated: 1789700916000
hash: CB87tYJg
---

# CurrentArray (属性)

如果指定单元格属于数组，则返回一个 **[Range](./../../Range/Range 对象.html)** 对象，该对象表示整个数组。只读。

## 示例

```javascript
/*假定工作表 Sheet1 的活动单元格为 A1，且该单元格属于包含单元格区域 A1:A10 的数组，那么本示例将选定工作表 Sheet1 上的单元格区域 A1:A10。*/
function test() {
    ActiveCell.CurrentArray.Select()
}
```

```javascript
/*本示例给 A1:C3 单元格设置数组公式，并显示 B2 单元格所包含的数组的整个地址*/
function test() {
    Range("A1:C3").FormulaArray = "=C3"
    console.log(Range("B2").CurrentArray.Address())
}
```
