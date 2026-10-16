---
title: "MarginRight (属性)"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/TextFrame2/%E5%B1%9E%E6%80%A7/MarginRight%20%E5%B1%9E%E6%80%A7.html
updated: 1789700916000
hash: NAsEjs5p
---

# MarginRight (属性)

返回或设置从文本框右边缘到包含文本的形状的内接矩形右边缘的距离（以磅为单位）。可读/写 **Single**类型。

## 示例

```javascript
/*本示例在第一张工作表中添加一个平行四边形，向平行四边形中添加文本，然后设置形状文本框右边距为 60 磅。*/
function test() {
    let textFrame2 = Application.Worksheets.Item(1).Shapes.AddShape(msoShapeParallelogram, 0, 0, 250, 140).TextFrame2
    textFrame2.TextRange.Text = "这是个平行四边形"
    textFrame2.MarginRight = 60
}
```

```javascript
/*本示例在活动工作表中添加一个椭圆，向椭圆中添加文本，并显示文本框架的右边距。*/
function test() {
    let textFrame2 = ActiveSheet.Shapes.AddShape(msoShapeOval, 100, 100, 200, 100).TextFrame2
    textFrame2.TextRange.Text = "这是个椭圆"
    console.log(textFrame2.MarginRight)
}
```
