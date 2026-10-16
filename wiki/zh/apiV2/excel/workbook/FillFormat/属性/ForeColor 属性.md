---
title: "ForeColor (属性)"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/FillFormat/%E5%B1%9E%E6%80%A7/ForeColor%20%E5%B1%9E%E6%80%A7.html
updated: 1789700916000
hash: CgO7hU_n
---

# ForeColor (属性)

返回或设置一个 **[ColorFormat](./../../ColorFormat/ColorFormat 对象.html)** 对象，它代表指定的前景填充色或纯色。

## 示例

```javascript
/*本示例在活动工作表中添加一个矩形，然后将矩形填充的前景色设置为红色。*/
function test() {
    let fillFormat = ActiveSheet.Shapes.AddShape(msoShapeRectangle, 100, 100, 100, 50).Fill
    fillFormat.ForeColor.RGB = RGB(255, 0, 0)
}
```

```javascript
/*本示例将第一张工作表中第二个形状的填充前景色的发光度设置为0.5。*/
function test() {
    let fillFormat = Application.Worksheets.Item(1).Shapes.Item(2).Fill
    fillFormat.ForeColor.Brightness = 0.5
}
```
