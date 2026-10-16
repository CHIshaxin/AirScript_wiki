---
title: "IndentLevel (属性)"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/Range/%E5%B1%9E%E6%80%A7/IndentLevel%20%E5%B1%9E%E6%80%A7.html
updated: 1789700916000
hash: Ql1LGIWu
---

# IndentLevel (属性)

返回或设置一个 **Variant** 值，它代表单元格或单元格区域的缩进量。可为 0 到 15 之间的整数。

## 说明

使用此属性将缩进量设为小于 0 或者大于 15 的数字，将导致发生错误。

## 示例

```javascript
/*此示例将单元格 A10 的缩进量增加至 15。*/
function test() {
    Range("A10").IndentLevel = 15
}
```

```javascript
/*本示例将 A1:C1 单元格区域的缩进量设置为 10，并显示 A1 单元格的缩进量*/
function test() {
    Range("A1:C1").IndentLevel = 10
    console.log(Range("A1").IndentLevel)
}
```
