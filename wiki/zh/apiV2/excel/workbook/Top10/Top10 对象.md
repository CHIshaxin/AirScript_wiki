---
title: "Top10 (对象)"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/Top10/Top10%20%E5%AF%B9%E8%B1%A1.html
updated: 1789700916000
hash: C_9AbhSm
---

# Top10 (对象)

代表条件格式规则的前十项。通过对某一区域应用颜色，有助于查看相对于其他单元格的单元格的值。

## 说明

所有条件格式设置对象都包含在 **[FormatConditions](./../FormatConditions/FormatConditions 对象.html)** 集合对象中，该对象是 **[Range](./../Range/Range 对象.html)** 集合的子对象。

可以使用 **[FormatConditions](./../FormatConditions/FormatConditions 对象.html)** 集合的 **Add** 或 **AddTop10** 方法创建前 10 个格式规则。

## 示例

```javascript
/*本示例通过条件格式规则生成一个动态数据集并对前 10 个值应用颜色。*/
function test() {
    //Building data
    Application.Range("A1").Value2 = "Name"
    Application.Range("B1").Value2 = "Number"
    Application.Range("A2").Value2 = "Agent1"
    Application.Range("A2").AutoFill(Application.Range("A2:A26"), xlFillDefault)
    Application.Range("B2:B26").FormulaArray = "=INT(RAND()*101)"
    Application.Range("B2:B26").Select()

    //Applying Conditional Formatting Top 10
    Application.Selection.FormatConditions.AddTop10()
    Application.Selection.FormatConditions.Item(Application.Selection.FormatConditions.Count).SetFirstPriority()
    let top = Application.Selection.FormatConditions.Item(1)
    top.TopBottom = xlTop10Top
    top.Rank = 10
    top.Percent = false

    //Applying color fill
    let font = Application.Selection.FormatConditions.Item(1).Font
    font.Color = RGB(0, 155, 115)
    font.TintAndShade = 0
    let interior = Application.Selection.FormatConditions.Item(1).Interior
    interior.PatternColorIndex = xlAutomatic
    interior.Color = RGB(5, 185, 115)
    interior.TintAndShade = 0
}
```

```javascript
/*本示例设置第一张工作表上区域 E1:E10 中第一个（Top10）条件格式的TopBottom属性，并将该条件格式设置为按百分比值确定排位，然后设置该条件格式的排位值的百分比。*/
function test() {
    let top = Worksheets.Item(1).Range("E1:E10").FormatConditions.Item(1)
    top.TopBottom = xlTop10Bottom
    top.Percent = true
    top.Rank = 20
}
```
