---
title: "MarginBottom (属性)"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/TextFrame/%E5%B1%9E%E6%80%A7/MarginBottom%20%E5%B1%9E%E6%80%A7.html
updated: 1789700916000
hash: DVxzi4b9
---

# MarginBottom (属性)

以磅为单位返回或设置从文本框底端到包含文本的形状中内接矩形底端的距离。可读/写。**Single** 类型。

## 示例

```javascript
/*本示例在第一张工作表中添加一个矩形，向矩形中添加文本，然后设置文本框架的边距。*/
function test() {
    let TextFrame = Application.Worksheets.Item(1).Shapes.AddShape(msoShapeRectangle, 0, 0, 250, 140).TextFrame
    TextFrame.Characters().Text = "Here is some test text"
    TextFrame.MarginBottom = 0
    TextFrame.MarginLeft = 100
    TextFrame.MarginRight = 0
    TextFrame.MarginTop = 20
}
```

```javascript
/*本示例在活动工作表中添加一个椭圆，向椭圆中添加文本，然后设置文本框架的底部边距。*/
function test() {
    let TextFrame = ActiveSheet.Shapes.AddShape(msoShapeOval, 100, 100, 200, 100).TextFrame
    TextFrame.Characters().Text = "这是个椭圆"
    TextFrame.MarginBottom = 10
}
```
