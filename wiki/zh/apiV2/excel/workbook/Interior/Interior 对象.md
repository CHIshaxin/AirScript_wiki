---
title: "Interior (对象)"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/Interior/Interior%20%E5%AF%B9%E8%B1%A1.html
updated: 1789700916000
hash: D3RXX7tF
---

# Interior (对象)

代表一个对象的内部。

## 说明

使用 **Interior** 属性可返回 **[Interior](./../Interior/Interior 对象.html)** 对象。

## 示例

```javascript
/*此示例将 Sheet1 中 A1 单元格内部设置为红色。*/
function test() {
    Application.Worksheets.Item("Sheet1").Range("A1").Interior.ColorIndex = 3
}
```

```javascript
/*本示例在活动工作表中 A1:C4 区域单元格的内部添加十字线图案。*/
function test() {
    ActiveSheet.Range("A1:C4").Interior.Pattern = xlPatternCrissCross
}
```
