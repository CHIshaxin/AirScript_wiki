---
title: "ColorScale (对象)"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/ColorScale/ColorScale%20%E5%AF%B9%E8%B1%A1.html
updated: 1789700916000
hash: -qHaSaWs
---

# ColorScale (对象)

代表色阶条件格式规则。

## 说明

所有条件格式对象均包含在 **[FormatConditions](./../FormatConditions/FormatConditions 对象.html)** 集合对象中，该集合对象是 **[Range](./../Range/Range 对象.html)** 集合的子项。您可以使用 **[FormatConditions](./../FormatConditions/FormatConditions 对象.html)** 集合的 **Add** 或 **AddColorScale** 方法创建色阶格式规则。

色阶是直观的参照，可以帮助您了解数据的分布和变化。您可以对数据范围、表中的数据或数据透视表中的数据应用双色或三色色阶。对于双色色阶条件格式，您可以将值、类型和颜色分配给范围的最小和最大阈值。三色色阶还具有中点阈值。

通过设置 **[ColorScaleCriteria](./../ColorScaleCriteria/ColorScaleCriteria 对象.html)** 对象的属性，可以确定其中的每个阈值。 **[ColorScaleCriteria](./../ColorScaleCriteria/ColorScaleCriteria 对象.html)** 对象是 **[ColorScale](./../ColorScale/ColorScale 对象.html)** 对象的子项，也是色阶的所有 **ColorScaleCriterion** 对象的集合。

## 示例

```javascript
/*本示例创建了一个数字范围，然后将双色色阶条件格式规则应用于该范围。然后指定最小阈值的颜色为红色，最大阈值的颜色为蓝色。*/
function test() {
    // Fill cells with sample data from 1 to 10
    ActiveSheet.Range("C1").Value2 = 1
    ActiveSheet.Range("C2").Value2 = 2
    ActiveSheet.Range("C1:C2").AutoFill(Range("C1:C10"))
    Range("C1:C10").Select()

    // Create a two-color ColorScale object for the created sample data range
    let colorScale = Selection.FormatConditions.AddColorScale(2)

    // Set the minimum threshold to red and maximum threshold to blue
    colorScale.ColorScaleCriteria(1).FormatColor.Color = RGB(255, 0, 0)
    colorScale.ColorScaleCriteria(2).FormatColor.Color = RGB(0, 0, 255)
}
```

```javascript
/*本示例设置活动工作表上单元格区域 A1:A10 的第一个色阶条件格式所应用于的单元格区域，并设置该条件格式的第一个阈值条件的颜色和亮度。*/
function test() {
    let colorScale = ActiveSheet.Range("A1:A10").FormatConditions.Item(1)
    colorScale.ModifyAppliesToRange(Range("A7:A8"))
    colorScale.ColorScaleCriteria.Item(1).FormatColor.ColorIndex = 7
    colorScale.ColorScaleCriteria.Item(1).FormatColor.TintAndShade = 0.5
}
```
