---
title: "ClearComments (方法)"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/Range/%E6%96%B9%E6%B3%95/ClearComments%20%E6%96%B9%E6%B3%95.html
updated: 1789700916000
hash: Dd8qZFtp
---

# ClearComments (方法)

清除指定区域的所有单元格批注。

## 示例

```javascript
/*本示例清除 E5 单元格的所有批注。*/
function test() {
    Application.Worksheets.Item(1).Range("E5").ClearComments()
}
```

```javascript
/*本示例设置 A1 单元格的批注为 a1，A2 单元格的批注为 a2，再清除 A1 单元格的批注*/
function test() {
    Range("A1").AddComment("a1")
    Range("A2").AddComment("a2")
    Range("A1").ClearComments()
}
```
