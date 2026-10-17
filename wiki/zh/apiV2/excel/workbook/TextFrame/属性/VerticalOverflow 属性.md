---
title: "VerticalOverflow (属性)"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/TextFrame/%E5%B1%9E%E6%80%A7/VerticalOverflow%20%E5%B1%9E%E6%80%A7.html
updated: 1789700916000
hash: BqJEiFRS
---

# VerticalOverflow (属性)

返回或设置指定对象的垂直溢出设置。可读/写。

## 说明

此属性只在 **AutoSize** 属性为 **False** 时有效。

## 返回值

**[XlOartVerticalOverflow](./../../枚举/XlOartVerticalOverflow 枚举.html)** 值

## 示例

```javascript
/*本示例在活动工作表中添加一个文本框，然后将文本框的垂直溢出设置为允许文本在垂直方向溢出文本框。*/
function test() {
    let textFrame = ActiveSheet.Shapes.AddTextbox(msoTextOrientationHorizontal, 100, 100, 200, 100).TextFrame
    textFrame.AutoSize = false
    textFrame.VerticalOverflow = xlOartVerticalOverflowOverflow
}
```

```javascript
/*本示例在第一张工作表中添加一个文本框，并显示文本框的垂直溢出是否设置成隐藏垂直方向溢出文本框的文本。*/
function test() {
    let textFrame = Application.Worksheets.Item(1).Shapes.AddTextbox(msoTextOrientationHorizontal, 100, 100, 200, 100).TextFrame
    console.log(textFrame.VerticalOverflow == xlOartVerticalOverflowClip)
}
```
