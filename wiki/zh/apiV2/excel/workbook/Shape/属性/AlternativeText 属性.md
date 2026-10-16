---
title: "AlternativeText (属性)"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/Shape/%E5%B1%9E%E6%80%A7/AlternativeText%20%E5%B1%9E%E6%80%A7.html
updated: 1789700916000
hash: CkadvTV2
---

# AlternativeText (属性)

返回或设置一个当 **[Shape](./../../Shape/Shape 对象.html)** 对象保存为网页时，该对象的描述性（可选）文本字符串。**String** 型，可读写。

## 说明

可选文字可以显示在 Web 浏览器中某个形状的图像位置上，或者当鼠标指针停留在图像上时，这些文本将直接显示在图像之上（在支持这些功能的浏览器中）。

## 示例

```javascript
/*本示例将活动工作表中第一个形状的可选文字设置为对该形状的说明。*/
function test() {
    let shape = ActiveSheet.Shapes.Item(1)
    shape.AlternativeText = "Concentric circles"
}
```

```javascript
/*本示例在第一张工作表中添加一个矩形，并将可选文字设置为对该形状的说明。*/
function test() {
    let shape = Worksheets(1).Shapes.AddShape(msoShapeRectangle, 300, 300, 100, 150)
    shape.AlternativeText = "这是一个矩形"
}
```
