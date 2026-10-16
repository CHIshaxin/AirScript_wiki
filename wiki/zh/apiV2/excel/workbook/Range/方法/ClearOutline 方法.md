---
title: "ClearOutline (方法)"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/Range/%E6%96%B9%E6%B3%95/ClearOutline%20%E6%96%B9%E6%B3%95.html
updated: 1789700916000
hash: BmNDZ7zX
---

# ClearOutline (方法)

清除指定区域的分级显示。

## 返回值

Variant

## 示例

```javascript
/*本示例清除工作表 Sheet1 上单元格区域 A1:G37 的分级显示。*/
function test() {
    Worksheets.Item("Sheet1").Range("A1:G37").ClearOutline()
}
```

```javascript
/*本示例清除 A1 单元格所在区域的分级显示，并重新创建分级显示*/
function test() {
    Range("A1").ClearOutline()
    console.log("已清除分级显示，关闭弹框后重建分级显示。")
    Range("A1").AutoOutline()
}
```
