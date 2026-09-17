---
title: "TextFrame2 属性"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/ShapeRange/%E5%B1%9E%E6%80%A7/TextFrame2%20%E5%B1%9E%E6%80%A7.html
---
# TextFrame2 (属性)

返回一个 **[TextFrame2](./../../TextFrame2/TextFrame2 对象.html)** 对象，该对象包含指定形状区域的文本格式。只读。

## 示例

```javascript
/*本示例先在第一张工作表中创建shpRange对象，将文本框内的字号设置为14。*/
function test() {
    let shpRange = Application.Worksheets.Item(1).Shapes.Range([1, 2])
    shpRange.TextFrame2.TextRange.Font.Size = 14
}
```

```javascript
/*本示例重新设置活动工作表的所有文本框的边距。*/
function test() {
    let shapes = ActiveSheet.Shapes
    shapes.SelectAll()
    Selection.ShapeRange.TextFrame2.MarginBottom = 15
    Selection.ShapeRange.TextFrame2.MarginLeft = 15
    Selection.ShapeRange.TextFrame2.MarginRight = 15
    Selection.ShapeRange.TextFrame2.MarginTop = 15
}
```
