---
title: "ShrinkToFit (属性)"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/Range/%E5%B1%9E%E6%80%A7/ShrinkToFit%20%E5%B1%9E%E6%80%A7.html
updated: 1789700916000
hash: BuchRDzj
---

# ShrinkToFit (属性)

返回或设置一个 **Variant** 值。

## 说明

如果文本自动收缩以适应可用列宽，此属性将返回 **True**；如果没有将指定区域中所有单元格的这一属性设为相同的值，则返回 **Null**。

## 示例

```javascript
/*此示例使第一行中的文本自动收缩为适当尺寸以适应可用列宽。*/
function test() {
    Rows.Item(1).ShrinkToFit = true
}
```

```javascript
/*本示例显示 A1:B3 单元格区域的文本是否自动收缩，然后将 A1 单元格设置文本自动收缩，再显示*/
function test() {
    let range = Range("A1:B3")
    console.log(range.ShrinkToFit) // false
    Range("A1").ShrinkToFit = true
    console.log(range.ShrinkToFit) // null
}
```
