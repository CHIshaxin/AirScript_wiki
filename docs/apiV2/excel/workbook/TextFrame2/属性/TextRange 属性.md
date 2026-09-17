---
title: "TextRange 属性"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/TextFrame2/%E5%B1%9E%E6%80%A7/TextRange%20%E5%B1%9E%E6%80%A7.html
---
# TextRange (属性)

返回 **TextRange2** 对象，该对象代表对象中的文本。只读

## 示例

```javascript
/*本示例在第一张工作表中添加一个矩形，向矩形中添加文本，然后显示文本的长度。*/
function test() {
    let textFrame2 = Application.Worksheets.Item(1).Shapes.AddShape(msoShapeRectangle, 0, 0, 250, 140).TextFrame2
    textFrame2.TextRange.Text = "这是个矩形"
    console.log(textFrame2.TextRange.Length)
}
```

```javascript
/*本示例在活动工作表中添加一个椭圆，向椭圆中添加文本，并显示文本框架中 TextRange2 集合中的项数和文本字体。*/
function test() {
    let textFrame2 = ActiveSheet.Shapes.AddShape(msoShapeOval, 100, 100, 200, 100).TextFrame2
    textFrame2.TextRange.Text = "这是个椭圆"
    console.log(textFrame2.TextRange.Count)
    console.log(textFrame2.TextRange.Font.Name)
}
```
