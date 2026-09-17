---
title: "ReadingOrder 属性"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/Range/%E5%B1%9E%E6%80%A7/ReadingOrder%20%E5%B1%9E%E6%80%A7.html
---
# ReadingOrder (属性)

返回或设置指定对象的阅读次序。可为以下常量之一：**xlRTL**（从右到左）、**xlLTR**（从左到右）或 **xlContext**。**Long** 类型，可读写。

## 示例

```javascript
/*本示例将 A1 单元格的阅读次序设置为根据上下文决定阅读次序*/
function test() {
    Range("A1").ReadingOrder = xlContext
}
```

```javascript
/*本示例选中 C4 单元格，将选中区域的阅读次序设置为从右到左阅读，将 E 列的阅读次序设置为从左到右阅读。*/
function test() {
    Range("C4").Select()
    Selection.ReadingOrder = xlRTL
    Columns.Item("E").ReadingOrder = xlLTR
}
```
