---
title: "VerticalAlignment 属性"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/TextFrame/%E5%B1%9E%E6%80%A7/VerticalAlignment%20%E5%B1%9E%E6%80%A7.html
---
# VerticalAlignment (属性)

返回或设置一个 **XlVAlign**值，它代表指定对象的垂直对齐方式。

## 说明

某些常量可能不可用，这取决于所选择或安装的语言支持（例如，美国英语）。

## 示例

```javascript
/*本示例在第一张工作表中添加一个平行四边形，向平行四边形中添加文本，并将垂直对齐方式设置为顶部对齐。*/
function test() {
    let textFrame = Application.Worksheets.Item(1).Shapes.AddShape(msoShapeParallelogram, 0, 0, 250, 140).TextFrame
    textFrame.Characters().Text = "这里是示例文本"
    textFrame.VerticalAlignment = xlVAlignTop
}
```

```javascript
/*本示例在活动工作表中添加一个文本框，然后显示文本框的垂直对齐方式是否为向下对齐 。*/
function test() {
    let textFrame = ActiveSheet.Shapes.AddTextbox(msoTextOrientationHorizontal, 100, 100, 200, 100).TextFrame
    console.log(textFrame.VerticalAlignment == xlVAlignBottom)
}
```
