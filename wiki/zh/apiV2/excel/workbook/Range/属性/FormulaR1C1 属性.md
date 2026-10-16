---
title: "FormulaR1C1 (属性)"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/Range/%E5%B1%9E%E6%80%A7/FormulaR1C1%20%E5%B1%9E%E6%80%A7.html
updated: 1789700916000
hash: BZ7FbCU1
---

# FormulaR1C1 (属性)

返回或设置指定对象的公式，使用宏语言 R1C1 格式符号表示。**Variant** 型，可读写。

## 说明

如果单元格包含一个常量，此属性返回该常量。如果单元格为空，此属性返回一个空字符串。如果单元格包含公式，此属性将该公式作为字符串返回，所用格式与在编辑栏（包括等号）中显示时的格式相同。

如果将单元格的格式的值或公式设为日期类型，ET 将检查该单元格的格式是否符合某个日期或时间数组格式，如果不符合，将采用默认的短日期数字格式。

如果指定区域是一维或二维区域，则可将公式指定为相同维数的数组。同样，也可在数组中使用公式。

对多重单元格区域设置公式，则该区域中所有单元格都用此公式填充。

## 示例

```javascript
/*此示例为 Sheet1 的 B1 单元格设置公式。*/
function test() {
    Worksheets.Item("Sheet1").Range("B1").FormulaR1C1 = "=SQRT(R1C1)"
}
```

```javascript
/*本示例为 A1 单元格设置公式 =R1C2+R1C3，并显示*/
function test() {
    Range("A1").FormulaR1C1 = '=R1C2+R1C3'
    console.log(Range("A1").FormulaR1C1)
}
```
