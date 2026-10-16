---
title: "TextFrame (属性)"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/ShapeRange/%E5%B1%9E%E6%80%A7/TextFrame%20%E5%B1%9E%E6%80%A7.html
updated: 1789700916000
hash: loBNte1q
---

# TextFrame (属性)

返回一个 **[TextFrame](./../../TextFrame/TextFrame 对象.html)** 对象，它包含指定形状的对齐方式和定位属性。

## 示例

```javascript
/*本示例先在第一张工作表创建shpRange对象，将中文本框内的文本设置为两端对齐。*/
function test() {
    let shpRange = Application.Worksheets.Item(1).Shapes.Range([1, 2])
    shpRange.TextFrame.HorizontalAlignment = xlHAlignJustify
}
```

```javascript
/*本示例重新设置活动工作表的所有文本框的边距。*/
function test() {
    let shapes = ActiveSheet.Shapes
    shapes.SelectAll()
    Selection.ShapeRange.TextFrame.MarginBottom = 10
    Selection.ShapeRange.TextFrame.MarginLeft = 10
    Selection.ShapeRange.TextFrame.MarginRight = 10
    Selection.ShapeRange.TextFrame.MarginTop = 10
}
```
