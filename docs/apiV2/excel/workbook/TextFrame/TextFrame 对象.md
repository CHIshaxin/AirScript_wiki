---
title: "TextFrame 对象"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/TextFrame/TextFrame%20%E5%AF%B9%E8%B1%A1.html
---
# TextFrame (对象)

代表 **[Shape](./../Shape/Shape 对象.html)** 对象中的文本框架。包含文本框架中的文本以及控制文本框架的对齐和定位的属性和方法。

## 说明

使用 **TextFrame** 属性可返回一个 **[TextFrame](./../TextFrame/TextFrame 对象.html)** 对象。

## 示例

```javascript
/*本示例在第一张工作表中添加一个矩形，向矩形中添加文本，然后设置文本框架的边距。*/
function test() {
    let textFrame = Application.Worksheets.Item(1).Shapes.AddShape(msoShapeRectangle, 0, 0, 250, 140).TextFrame
    textFrame.Characters().Text = "Here is some test text"
    textFrame.MarginBottom = 10
    textFrame.MarginLeft = 10
    textFrame.MarginRight = 10
    textFrame.MarginTop = 10
}
```

```javascript
/*本示例在活动工作表中添加一个椭圆，向椭圆中添加文本，并将指定文本颜色改为绿色。*/
function test() {
    let textFrame = ActiveSheet.Shapes.AddShape(msoShapeOval, 100, 100, 200, 100).TextFrame
    textFrame.Characters().Text = "这是示例文本"
    textFrame.Characters(3, 2).Font.Color = RGB(0, 255, 0)
}
```
