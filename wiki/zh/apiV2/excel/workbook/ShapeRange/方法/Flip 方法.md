---
title: "Flip (方法)"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/ShapeRange/%E6%96%B9%E6%B3%95/Flip%20%E6%96%B9%E6%B3%95.html
updated: 1789700916000
hash: CnCsnlBC
---

# Flip (方法)

绕指定形状的水平或垂直对称轴翻转该形状。

## 参数

| 属性 | 数据类型 | 必填 | 说明 |
| --- | --- | --- | --- |
| FlipCmd | Kso.MsoFlipCmd | 必填 | 指定形状是水平翻转还是垂直翻转。 |

## 示例

```javascript
/*本示例首先向第一张工作表添加一个三角形，然后通过新建的ShapeRange对象设置成红色并垂直翻转。*/
function test() {
    Application.Worksheets.Item(1).Shapes.AddShape(msoShapeRightTriangle, 10, 10, 50, 50).Name = "三角形"
    let shpRange = Application.Worksheets.Item(1).Shapes.Range(["三角形"])
    shpRange.Fill.ForeColor.RGB = RGB(255, 0, 0)
    shpRange.Flip(msoFlipVertical)
}
```

```javascript
/*本示例将活动工作表的所有形状水平翻转。*/
function test() {
    let shapes = ActiveSheet.Shapes
    shapes.SelectAll()
    Selection.ShapeRange.Flip(msoFlipHorizontal)
}
```
