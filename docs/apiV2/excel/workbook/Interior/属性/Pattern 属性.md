---
title: "Pattern 属性"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/Interior/%E5%B1%9E%E6%80%A7/Pattern%20%E5%B1%9E%E6%80%A7.html
---
# Pattern (属性)

返回或设置一个包含 **xlPattern** 常量的 **Variant** 值，它代表内部图案。

## 示例

```javascript
/*本示例在 Sheet1 中 A1 单元格的内部添加十字图案。*/
function test() {
    Application.Worksheets.Item("Sheet1").Range("A1").Interior.Pattern = xlPatternCrissCross
}
```

```javascript
/*本示例在活动工作表中 A1:C4 单元格的内部添加网格图案。*/
function test() {
    ActiveSheet.Range("A1:C4").Interior.Pattern = xlPatternGrid
}
```
