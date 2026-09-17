---
title: "FormulaLocal 属性"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/Range/%E5%B1%9E%E6%80%A7/FormulaLocal%20%E5%B1%9E%E6%80%A7.html
---
# FormulaLocal (属性)

返回或设置指定对象的公式，使用用户语言 A1 格式引用。**Variant** 型，可读写。

## 说明

如果指定单元格包含常量，此属性返回的就是该常量。如果单元格为空，此属性返回一个空字符串。如果单元格包含公式，此属性将该公式作为字符串返回，所用格式与在编辑栏（包括等号）中显示时的格式相同。

如果将单元格的格式的值或公式设为日期类型，ET 将检查该单元格的格式是否符合某个日期或时间数组格式，如果不符合，将采用默认的短日期数字格式。

如果指定区域是一维或二维区域，则可将公式指定为相同维数的数组。同样，也可在数组中使用公式。

对多重单元格区域设置公式，则该区域中所有单元格都用此公式填充。

## 示例

```javascript
/*假定使用的是美国英语版 WPS 表格，并往第一张工作表的 A11 单元格中输入了公式“=SUM(A1:A10)”。如果在一台运行德文版 WPS 表格 的计算机上打开该工作表，并运行下例，该示例将在消息框中显示公式“=SUMME(A1:A10)”。*/
function test() {
    Worksheets.Item(1).Range("A11").Formula = '=SUM(A1:A10)'
    console.log(Worksheets.Item(1).Range("A11").FormulaLocal)
}
```

```javascript
/*本示例将 A1 单元格的公式设置为 =TODAY()，并显示 A1 单元格的公式*/
function test() {
    Range("A1").Formula = "=TODAY()"
    console.log(Range("A11").FormulaLocal)
}
```
