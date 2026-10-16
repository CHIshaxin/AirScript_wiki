---
title: "AutoOutline (方法)"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/Range/%E6%96%B9%E6%B3%95/AutoOutline%20%E6%96%B9%E6%B3%95.html
updated: 1789700916000
hash: 7lZ3SeJ4
---

# AutoOutline (方法)

自动为指定区域创建分级显示。如果区域为单个单元格，ET 将创建整个工作表的分级显示。新分级显示将取代所有的分级显示。

## 返回值

Variant

## 示例

```javascript
/*本示例创建工作表 Sheet1 上 A1 单元格所在区域的分级显示。该区域必须包含汇总行或汇总列。*/
function test() {
    Worksheets.Item("Sheet1").Range("A1").AutoOutline()
}
```

```javascript
/*本示例选中 C1 单元格，并创建选中单元格所在区域的分级显示。*/
function test() {
    Range("C1").Select()
    Selection.AutoOutline()
}
```
