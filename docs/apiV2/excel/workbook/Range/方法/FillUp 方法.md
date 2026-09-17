---
title: "FillUp 方法"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/Range/%E6%96%B9%E6%B3%95/FillUp%20%E6%96%B9%E6%B3%95.html
---
# FillUp (方法)

从指定区域的底部单元格开始向上填充，直至该区域的顶部。区域中尾行单元格的内容和格式将复制到区域中其他行内。

## 返回值

Variant

## 示例

```javascript
/*本示例基于工作表 Sheet1 上单元格 A10 的内容，填充单元格区域 A1:A10。*/
function test() {
    Worksheets.Item("Sheet1").Range("A1:A10").FillUp()
}
```

```javascript
/*本示例将 A5:B5 单元格区域值设为 5，B5 单元格字体设置为斜体，将 A1:B5 单元格区域根据最下侧单元格向上填充*/
function test() {
    Range("A5:B5").Value2 = 5
    Range("B5").Font.Italic = true
    Range("A1:B5").FillUp()
}
```
