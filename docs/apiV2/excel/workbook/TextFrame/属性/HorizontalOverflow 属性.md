---
title: "HorizontalOverflow 属性"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/TextFrame/%E5%B1%9E%E6%80%A7/HorizontalOverflow%20%E5%B1%9E%E6%80%A7.html
---
# HorizontalOverflow (属性)

返回或设置指定对象的水平溢出设置。可读/写。

## 说明

此属性只在 **WordWrap** 属性为 **msoFalse** (0) 时有效。

## 返回值

**XlOartHorizontalOverflow值**

## 示例

```javascript
/*本示例在活动工作表中添加一个文本框，然后将文本框的水平溢出设置为允许文本在水平方向溢出文本框。*/
function test() {
    let textFrame = ActiveSheet.Shapes.AddTextbox(msoTextOrientationHorizontal, 100, 100, 200, 100).TextFrame
    textFrame.HorizontalOverflow = xlOartHorizontalOverflowOverflow
}
```

```javascript
/*本示例在第一张工作表中添加一个文本框，并显示文本框的水平溢出是否设置成隐藏水平方向溢出文本框的文本。*/
function test() {
    let textFrame = Application.Worksheets.Item(1).Shapes.AddTextbox(msoTextOrientationHorizontal, 100, 100, 200, 100).TextFrame
    console.log(textFrame.HorizontalOverflow == xlOartHorizontalOverflowClip)
}
```
