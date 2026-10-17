---
title: "Orientation (属性)"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/TextFrame/%E5%B1%9E%E6%80%A7/Orientation%20%E5%B1%9E%E6%80%A7.html
updated: 1789700916000
hash: Hzmekf4b
---

# Orientation (属性)

返回或设置一个 **Long** 值，它代表文本框的方向。

## 说明

此属性值可设为 -90 到 90 度之间的整数值或以下 **MsoTextOrientation** 常量之一。

## 示例

```javascript
/*本示例在活动工作表中添加一个文本框，向文本框中添加文本，然后设置文本框的方向为纵向 。*/
function test() {
    let textFrame = ActiveSheet.Shapes.AddTextbox(msoTextOrientationHorizontal, 100, 100, 200, 100).TextFrame
    textFrame.Characters().Text = "示例文本"
    textFrame.Orientation = msoTextOrientationVertical
}
```

```javascript
/*本示例在第一张工作表中添加一个文本框，向文本框中添加文本，并显示文本框的方向是否是水平方向。*/
function test() {
    let textFrame = Application.Worksheets.Item(1).Shapes.AddTextbox(msoTextOrientationHorizontal, 100, 100, 200, 100).TextFrame
    textFrame.Characters().Text = "示例文本"
    console.log(textFrame.Orientation == msoTextOrientationHorizontal)
}
```
