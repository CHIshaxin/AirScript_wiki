---
title: "ShrinkToFit 属性"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/CellFormat/%E5%B1%9E%E6%80%A7/ShrinkToFit%20%E5%B1%9E%E6%80%A7.html
---
# ShrinkToFit (属性)

返回或设置一个 **Variant** 值。该值表示文本是否自动缩小以适应可用列宽

## 说明

如果文本自动收缩以适应可用列宽，此属性将返回 **True**；如果没有将指定区域中所有单元格的这一属性设为相同的值，则返回 **Null**。

## 示例

```javascript
/*此示例显示查找条件中文本是否自动收缩以适应可用的列宽。*/
function test() {
    console.log(Application.FindFormat.ShrinkToFit)
}
```

```javascript
/*此示例设置替换条件中文本为自动收缩以适应可用的列宽。*/
function test() {
    Application.ReplaceFormat.ShrinkToFit = true
}
```
