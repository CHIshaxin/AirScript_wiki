---
title: "Flip 方法"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/Shape/%E6%96%B9%E6%B3%95/Flip%20%E6%96%B9%E6%B3%95.html
---
# Flip (方法)

绕指定形状的水平或垂直对称轴翻转该形状。

## 参数

| 属性 | 数据类型 | 必填 | 说明 |
| --- | --- | --- | --- |
| FlipCmd | Kso.MsoFlipCmd | 必填 | 指定形状是水平翻转还是垂直翻转。 |

## 示例

```javascript
/*本示例首先向活动工作表添加一个三角形，然后设置为红色并垂直翻转。*/
function test() {
    let shape = Application.ActiveSheet.Shapes.AddShape(msoShapeRightTriangle, 10, 10, 50, 50)
    shape.Fill.ForeColor.RGB = RGB(255, 0, 0)
    shape.Flip(msoFlipVertical)
}
```

```javascript
/*本示例将第一张工作表中的第一个形状水平翻转。*/
function test() {
    let shape = ActiveSheet.Shapes.Item(1)
    shape.Flip(msoFlipHorizontal)
}
```
