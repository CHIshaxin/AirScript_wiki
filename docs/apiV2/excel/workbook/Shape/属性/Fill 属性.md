---
title: "Fill 属性"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/Shape/%E5%B1%9E%E6%80%A7/Fill%20%E5%B1%9E%E6%80%A7.html
---
# Fill (属性)

返回指定形状的 **[FillFormat](./../../FillFormat/FillFormat 对象.html)** 对象或指定图表的 **ChartFillFormat** 对象，这些对象包含形状或图表的填充格式属性。只读。

## 示例

```javascript
/*本示例在活动工作表中添加一个矩形，然后设置矩形填充的前景色、背景色和渐变。*/
function test() {
    let shape = ActiveSheet.Shapes.AddShape(msoShapeRectangle, 90, 90, 90, 50)
    shape.Fill.ForeColor.RGB = RGB(128, 0, 0)
    shape.Fill.BackColor.RGB = RGB(170, 170, 170)
    shape.Fill.TwoColorGradient(msoGradientHorizontal, 1)
}
```

```javascript
/*本示例判断如果活动工作表中第二个形状前景色为红色，则前景色设置为蓝色。*/
function test() {
    let shapes = ActiveSheet.Shapes
    if (shapes.Item(2).Fill.ForeColor.RGB == RGB(255, 0, 0)) {
        shapes.Item(2).Fill.ForeColor.RGB = RGB(0, 0, 255)
    }
}
```
