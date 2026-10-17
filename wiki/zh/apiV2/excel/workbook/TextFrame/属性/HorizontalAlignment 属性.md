---
title: "HorizontalAlignment (属性)"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/TextFrame/%E5%B1%9E%E6%80%A7/HorizontalAlignment%20%E5%B1%9E%E6%80%A7.html
updated: 1789700916000
hash: C9IJ1H9p
---

# HorizontalAlignment (属性)

返回或设置一个 **XlHAlign**值，它代表指定对象的水平对齐方式。

## 说明

某些常量可能不可用，这取决于所选择或安装的语言支持（例如，美国英语）。

## 示例

```javascript
/*本示例在第一张工作表中添加一个矩形，向矩形中添加文本，并将水平对齐方式设置为左对齐。*/
function test() {
    let textFrame = Application.Worksheets.Item(1).Shapes.AddShape(msoShapeRectangle, 0, 0, 250, 140).TextFrame
    textFrame.Characters().Text = "这里是示例文本"
    textFrame.HorizontalAlignment = xlHAlignLeft
}
```

```javascript
/*本示例在活动工作表中添加一个文本框，然后显示文本框的水平对齐方式是否为右对齐 。*/
function test() {
    let textFrame = ActiveSheet.Shapes.AddTextbox(msoTextOrientationHorizontal, 100, 100, 200, 100).TextFrame
    console.log(textFrame.HorizontalAlignment == xlHAlignRight)
}
```
