---
title: "AboveAverage 对象"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/AboveAverage/AboveAverage%20%E5%AF%B9%E8%B1%A1.html
---
# AboveAverage (对象)

代表条件格式规则的高于平均值的视图。对某一区域或选定内容应用颜色或填充有助于您查看与其他单元格相关的单元格的值。

## 说明

所有条件格式对象均包含在 **[FormatConditions](./../FormatConditions/FormatConditions 对象.html)** 集合对象中，该集合对象是 **[Range](./../Range/Range 对象.html)** 集合的子项。您可以使用 **[FormatConditions](./../FormatConditions/FormatConditions 对象.html)** 集合的 **Add** 或 **AddAboveAverage** 方法创建高于平均值格式规则。

## 示例

```javascript
/*本示例通过条件格式规则生成一个动态数据集并对高于平均值的值应用颜色。*/
function test() {
    //Building data for Melanie
    Range("A1").Value2 = "Name"
    Range("B1").Value2 = "Number"
    Range("A2").Value2 = "Melanie-1"
    Range("A2").AutoFill(Range("A2:A26"), xlFillDefault)
    Range("B2:B26").FormulaArray = "=INT(RAND()*101)"
    Range("B2:B26").Select()

    //Applying Conditional Formatting to items above the average.  Should appear green fill and dark green font.
    Selection.FormatConditions.AddAboveAverage()
    Selection.FormatConditions(Selection.FormatConditions.Count).SetFirstPriority()
    Selection.FormatConditions(1).AboveBelow = xlAboveAverage
    let font = Selection.FormatConditions(1).Font
    font.Color = RGB(0, 155, 115)
    font.TintAndShade = 0
    let interior = Selection.FormatConditions(1).Interior
    interior.PatternColorIndex = xlAutomatic
    interior.Color = RGB(5, 185, 115)
    interior.TintAndShade = 0
    console.log("Added an Above Average Conditional Format to Melanie's data.  Press F9 to update values.")
}
```

```javascript
/*本示例设置工作表 Sheet1 上区域 C1:C10 中第一个（AboveAverage）条件格式的AboveBelow属性，并设置该条件格式的内部颜色。*/
function test() {
    let aboveAverage = Application.Worksheets.Item("Sheet1").Range("C1:C10").FormatConditions.Item(1)
    aboveAverage.AboveBelow = xlEqualBelowAverage
    aboveAverage.Interior.ColorIndex = 10
}
```
