---
title: "FillDown (方法)"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/Range/%E6%96%B9%E6%B3%95/FillDown%20%E6%96%B9%E6%B3%95.html
updated: 1789700916000
hash: xDY4tEfU
---

# FillDown (方法)

从指定区域的顶部单元格开始向下填充，直至该区域的底部。区域中首行单元格的内容和格式将复制到区域中其他行内。

## 返回值

Variant

## 示例

```javascript
/*本示例基于单元格 A1 的内容填写工作表 Sheet1 的单元格区域 A1:A10。*/
function test() {
    Worksheets.Item("Sheet1").Range("A1:A10").FillDown()
}
```

```javascript
/*本示例将 A1 单元格字体加粗，A1:B1 单元格区域设置值为 2，基于顶部单元格内容填充单元格区域 A1:B5。*/
function test() {
    Range("A1").Font.Bold = true
    Range("A1:B1").Value2 = 2
    Range("A1:B5").FillDown()
}
```
