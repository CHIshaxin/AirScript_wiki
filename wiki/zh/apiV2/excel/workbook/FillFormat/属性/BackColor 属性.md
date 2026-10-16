---
title: "BackColor (属性)"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/FillFormat/%E5%B1%9E%E6%80%A7/BackColor%20%E5%B1%9E%E6%80%A7.html
updated: 1789700916000
hash: Citu9H-V
---

# BackColor (属性)

返回或设置一个 **[ColorFormat](./../../ColorFormat/ColorFormat 对象.html)** 对象，它代表指定的填充背景色。

## 示例

```javascript
/*本示例在活动工作表中添加一个矩形，然后将矩形填充的背景色设置为蓝色。*/
function test() {
    let fillFormat = ActiveSheet.Shapes.AddShape(msoShapeRectangle, 100, 100, 100, 50).Fill
    fillFormat.BackColor.RGB = RGB(0, 0, 255)
}
```

```javascript
/*本示例将第一张工作表中第二个形状的填充背景色的发光度设置为0.7。*/
function test() {
    let fillFormat = Application.Worksheets.Item(1).Shapes.Item(2).Fill
    fillFormat.BackColor.Brightness = 0.7
}
```
