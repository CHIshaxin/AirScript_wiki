---
title: "TextFrame2 对象"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/TextFrame2/TextFrame2%20%E5%AF%B9%E8%B1%A1.html
---
# TextFrame2 (对象)

代表 **Shape**、**ShapeRange** 或 **[ChartFormat](./../ChartFormat/ChartFormat 对象.html)** 对象的文本框。

## 说明

该对象包含文本框中的文本，还包含控制文本框对齐方式和位置的属性和方法。使用 **TextFrame2**属性可返回 **[TextFrame2](./../TextFrame2/TextFrame2 对象.html)** 对象。

## 示例

```javascript
/*本示例在第一张工作表中添加一个矩形，向矩形中添加文本，然后设置文本框的边距。*/
function test() {
    let textFrame2 = Application.Worksheets.Item(1).Shapes.AddShape(msoShapeRectangle, 0, 0, 250, 140).TextFrame2
    textFrame2.TextRange.Text = "Here is some test text"
    textFrame2.MarginBottom = 10
    textFrame2.MarginLeft = 10
    textFrame2.MarginRight = 10
    textFrame2.MarginTop = 10
}
```

```javascript
/*本示例判断如果活动工作表中第二个形状的文本框包含文本，则显示文本长度。*/
function test() {
    let textFrame2 = ActiveSheet.Shapes.Item(2).TextFrame2
    if (textFrame2.HasText == msoTrue) {
        console.log(textFrame2.TextRange.Length)
    }
}
```
