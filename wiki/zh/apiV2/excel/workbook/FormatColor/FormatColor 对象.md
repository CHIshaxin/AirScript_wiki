---
title: "FormatColor (对象)"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/FormatColor/FormatColor%20%E5%AF%B9%E8%B1%A1.html
updated: 1789700916000
hash: BSWj5D16
---

# FormatColor (对象)

代表为色阶条件格式阈值指定的填充色或数据条条件格式的条形颜色。

## 说明

您可以通过传递 **Color** 属性中的 RGB 值来选择颜色，或者通过使用 **ThemeColor** 属性在主题调色板中编制索引来指定颜色。

以下代码示例创建了一个数字范围，然后将双色色阶条件格式规则应用于该范围。然后通过在 **[ColorScaleCriteria](./../ColorScaleCriteria/ColorScaleCriteria 对象.html)** 集合中编制索引来设置单独的条件，从而指定最小阈值的颜色为红色，最大阈值的颜色为蓝色。

## 示例

```javascript
/*以下代码示例创建了一个数字范围，然后将双色色阶条件格式规则应用于该范围。然后，通过索引到 ColorScaleCriteria 集合以设置单个条件，将最小阈值的颜色分配给红色，并将最大阈值分配给蓝色。*/
function test() {
    //Fill cells with sample data from 1 to 10
    ActiveSheet.Range("C1").Value2 = 1
    ActiveSheet.Range("C2").Value2 = 2
    ActiveSheet.Range("C1:C2").AutoFill(Range("C1:C10"))

    Range("C1:C10").Select()

    //Create a two-color ColorScale object for the created sample data range
    let colorScale = Selection.FormatConditions.AddColorScale(2)

    //Set the minimum threshold to red and maximum threshold to blue
    colorScale.ColorScaleCriteria.Item(1).FormatColor.Color = RGB(255, 0, 0)
    colorScale.ColorScaleCriteria.Item(2).FormatColor.Color = RGB(0, 0, 255)
}
```

```javascript
/*本示例显示指定单元格区域的第一个色阶条件格式中第二个条件的最大阈值的主题颜色是否为xlThemeColorAccent1。*/
function test() {
    let formatCondition = ActiveSheet.Range("A1:A10").FormatConditions.Item(1)
    console.log(formatCondition.ColorScaleCriteria.Item(2).FormatColor.ThemeColor == xlThemeColorAccent1)
}
```
