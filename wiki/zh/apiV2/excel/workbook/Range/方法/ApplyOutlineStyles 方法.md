---
title: "ApplyOutlineStyles (方法)"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/Range/%E6%96%B9%E6%B3%95/ApplyOutlineStyles%20%E6%96%B9%E6%B3%95.html
updated: 1789700916000
hash: BLR3TZvT
---

# ApplyOutlineStyles (方法)

对指定区域应用分级显示样式。

## 返回值

Variant

## 示例

```javascript
/*下例对选定区域应用自动分级显示样式。该选定区域必须包括工作表上的整个分级显示区域。*/
function test() {
    Selection.ApplyOutlineStyles()
}
```

```javascript
/*本示例将 A1:C20 单元格区域应用分级显示样式*/
function test() {
    Range("A1:C20").ApplyOutlineStyles()
}
```
