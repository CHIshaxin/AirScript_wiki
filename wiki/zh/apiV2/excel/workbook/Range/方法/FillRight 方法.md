---
title: "FillRight (方法)"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/Range/%E6%96%B9%E6%B3%95/FillRight%20%E6%96%B9%E6%B3%95.html
updated: 1789700916000
hash: Dum4-405
---

# FillRight (方法)

从指定区域的最左边单元格开始向右填充。区域中最左列单元格的内容和格式将复制到区域中其他列内。

## 返回值

Variant

## 示例

```javascript
/*本示例基于工作表 Sheet1 上单元格 A1 的内容，填充单元格区域 A1:M1。*/
function test() {
    Worksheets.Item("Sheet1").Range("A1:M1").FillRight()
}
```

```javascript
/*本示例将 A1:A2 单元格区域的值设置为 2，A2 字体加粗，将 A1:G2 单元格区域根据最左侧单元格向右填充*/
function test() {
    Range("A1:A2").Value2 = 2
    Range("A2").Font.Bold = true
    Range("A1:G2").FillRight()
}
```
