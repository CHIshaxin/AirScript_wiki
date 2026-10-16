---
title: "UniqueValues 对象"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/UniqueValues/UniqueValues%20%E5%AF%B9%E8%B1%A1.html
---
# UniqueValues (对象)

**[UniqueValues](./../UniqueValues/UniqueValues 对象.html)** 对象使用 **DupeUnique** 属性返回或设置一个枚举，该枚举确定规则是查找区域中的重复值还是唯一值。

## 示例

```javascript
/*本示例通过条件格式规则生成一个动态数据集并对重复值应用颜色。*/
function test() {
    Application.Range("A1").Value2 = "Name"
    Application.Range("B1").Value2 = "Number"
    Application.Range("A2").Value2 = "Agent1"
    Application.Range("A2").AutoFill(Application.Range("A2:A26"), xlFillDefault)
    Application.Range("B2:B26").FormulaArray = "=INT(RAND()*101)"
    Application.Range("B2:B26").Select()

    Application.Selection.FormatConditions.AddUniqueValues()
    Application.Selection.FormatConditions.Item(Application.Selection.FormatConditions.Count).SetFirstPriority()
    Application.Selection.FormatConditions.Item(1).DupeUnique = xlDuplicate

    let font = Application.Selection.FormatConditions.Item(1).Font
    font.ColorIndex = 3
    font.TintAndShade = 0
    let interior = Application.Selection.FormatConditions.Item(1).Interior
    interior.PatternColorIndex = xlAutomatic
    interior.ColorIndex = 5
    interior.TintAndShade = 0
}
```

```javascript
/*本示例设置第一张工作表上单元格区域 C1:C10 的第一个（UniqueValues）条件格式所应用于的单元格区域，并设置该条件格式内部唯一值的颜色。*/
function test() {
    let uniqueValues = Application.Worksheets.Item(1).Range("C1:C10").FormatConditions.Item(1)
    uniqueValues.ModifyAppliesToRange(Range("C7:C10"))
    uniqueValues.DupeUnique = xlUnique
    uniqueValues.Interior.ColorIndex = 7
}
```
