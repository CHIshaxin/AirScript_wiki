---
title: "TextFrame 属性"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/Shape/%E5%B1%9E%E6%80%A7/TextFrame%20%E5%B1%9E%E6%80%A7.html
---
# TextFrame (属性)

返回一个 **[TextFrame](./../../TextFrame/TextFrame 对象.html)** 对象，它包含指定形状的对齐方式和定位属性。

## 示例

```javascript
/*本示例将第一张工作表中的第一个形状中文本框内的文本设置为两端对齐。*/
function test() {
    Application.Worksheets.Item(1).Shapes.Item(1).TextFrame.HorizontalAlignment = xlHAlignJustify
}
```

```javascript
/*下例向活动工作表中添加一个矩形，向矩形中添加文本，然后设置文本框的边距。*/
function test() {
    let textFrame = ActiveSheet.Shapes.AddShape(msoShapeRectangle, 0, 0, 250, 140).TextFrame
    textFrame.Characters().Text = "Here is some test text"
    textFrame.MarginBottom = 10
    textFrame.MarginLeft = 10
    textFrame.MarginRight = 10
    textFrame.MarginTop = 10
}
```
