---
title: "ColorStops (对象)"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/ColorStops/ColorStops%20%E5%AF%B9%E8%B1%A1.html
updated: 1789700916000
hash: BK1J3qN4
---

# ColorStops (对象)

指定的数据系列中所有 ColorStop 对象的集合。

## 说明

每个 **[ColorStop](./../ColorStop/ColorStop 对象.html)** 对象代表一个区域或选定内容中渐变填充的一个颜色光圈。

## 示例

```javascript
/*以下示例显示了具有线性渐变的颜色停止。*/
function test() {
    let interior = Selection.Interior
    interior.Pattern = xlPatternLinearGradient
    interior.Gradient.Degree = 90
    interior.Gradient.ColorStops.Clear()

    //adds stops after any have been cleared
    let colorStops1 = Selection.Interior.Gradient.ColorStops.Add(0)
    colorStops1.ThemeColor = xlThemeColorDark1
    colorStops1.TintAndShade = 0

    let colorStops2 = Selection.Interior.Gradient.ColorStops.Add(1)
    colorStops2.ThemeColor = xlThemeColorAccent1
    colorStops2.TintAndShade = 0
}
```

```javascript
/*以下示例显示了具有矩形渐变的颜色停止点。*/
function test() {
    let interior = Selection.Interior
    interior.Pattern = xlPatternRectangularGradient
    interior.Gradient.RectangleLeft = 0
    interior.Gradient.RectangleRight = 0
    interior.Gradient.RectangleTop = 0
    interior.Gradient.RectangleBottom = 0
    interior.Gradient.ColorStops.Clear()

    let colorStops1 = Selection.Interior.Gradient.ColorStops.Add(0)
    colorStops1.Color = 192
    colorStops1.TintAndShade = 0

    let colorStops2 = Selection.Interior.Gradient.ColorStops.Add(1)
    colorStops2.ThemeColor = xlThemeColorLight1
    colorStops2.TintAndShade = 0
}
```
