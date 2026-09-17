---
title: "ColorStop 对象"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/ColorStop/ColorStop%20%E5%AF%B9%E8%B1%A1.html
---
# ColorStop (对象)

表示某一区域或选定内容中渐变填充的光圈点。

## 说明

**[ColorStop](./../ColorStop/ColorStop 对象.html)** 对象可用来设置单元格填充属性，包括 **Color**、**ThemeColor** 和 **TintAndShade** 属性。

## 示例

```javascript
/*以下示例演示如何对 ColorStop 对象应用属性。*/
function test() {
    let interior = Application.Selection.Interior
    interior.Pattern = xlPatternLinearGradient
    interior.Gradient.Degree = 135
    interior.Gradient.ColorStops.Clear()

    let colorStops1 = Selection.Interior.Gradient.ColorStops.Add(0)
    colorStops1.ThemeColor = xlThemeColorDark1
    colorStops1.TintAndShade = 0

    let colorStops2 = Application.Selection.Interior.Gradient.ColorStops.Add(0.5)
    colorStops2.ThemeColor = xlThemeColorAccent1
    colorStops2.TintAndShade = 0

    let colorStops3 = Application.Selection.Interior.Gradient.ColorStops.Add(1)
    colorStops3.ThemeColor = xlThemeColorDark1
    colorStops3.TintAndShade = 0
}
```

```javascript
/*本示例显示当前选定区域的第一个ColorStop对象的位置。*/
function test() {
    Range("A1:A10").Select()
    let colorStop = Selection.Interior.Gradient.ColorStops.Item(1)
    console.log(colorStop.Position)
}
```
