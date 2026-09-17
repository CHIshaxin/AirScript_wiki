---
title: "ClearNotes 方法"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/Range/%E6%96%B9%E6%B3%95/ClearNotes%20%E6%96%B9%E6%B3%95.html
---
# ClearNotes (方法)

清除指定区域中所有单元格的批注和语音批注。

## 返回值

Variant

## 示例

```javascript
/*本示例清除工作表 Sheet1 上从 A 列到 C 列中的所有批注和语音批注。*/
function test() {
    Worksheets.Item("Sheet1").Columns.Item("A:C").ClearNotes()
}
```

```javascript
/*本示例设置 A1 单元格的批注为 a1，A2 单元格的批注为 a2，并清除 A1 单元格的所有批注和语音批注*/
function test() {
    Range("A1").AddComment("a1")
    Range("A2").AddComment("a2")
    Range("A1").ClearNotes()
}
```
