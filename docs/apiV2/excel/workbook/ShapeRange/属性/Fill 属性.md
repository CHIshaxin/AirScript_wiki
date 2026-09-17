---
title: "Fill 属性"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/ShapeRange/%E5%B1%9E%E6%80%A7/Fill%20%E5%B1%9E%E6%80%A7.html
---
# Fill (属性)

返回指定形状的 **[FillFormat](./../../FillFormat/FillFormat 对象.html)** 对象或指定图表的 **ChartFillFormat** 对象，这些对象包含形状或图表的填充格式属性。只读。

## 示例

```javascript
/*本示例为第一张工作表中新建的ShapeRange对象设置填充格式的前景色、背景色和渐变。*/
function test() {
    let shpRange = Application.Worksheets.Item(1).Shapes.Range([1, 2])
    shpRange.Fill.ForeColor.RGB = RGB(128, 0, 0)
    shpRange.Fill.BackColor.RGB = RGB(170, 170, 170)
    shpRange.Fill.TwoColorGradient(msoGradientHorizontal, 1)
}
```

```javascript
/*本示例将活动工作表的所有形状的前景色设置为红色。*/
function test() {
    let shapes = ActiveSheet.Shapes
    shapes.SelectAll()
    Selection.ShapeRange.Fill.ForeColor.RGB = RGB(255, 0, 0)
}
```
