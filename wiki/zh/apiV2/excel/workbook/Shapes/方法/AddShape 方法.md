---
title: "AddShape (方法)"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/Shapes/%E6%96%B9%E6%B3%95/AddShape%20%E6%96%B9%E6%B3%95.html
updated: 1789700916000
hash: D1tAwNtQ
---

# AddShape (方法)

返回一个 **[Shape](./../../Shape/Shape 对象.html)** 对象，该对象表示工作表中的新自选形状。

## 说明

要更改已添加的自选形状的类型，请设置 **AutoShapeType** 属性。

## 参数

| 属性 | 数据类型 | 必填 | 说明 |
| --- | --- | --- | --- |
| Type | Kso.MsoAutoShapeType | 必填 | 指定要创建的自选形状的类型。 |
| Left | float | 必填 | 自选形状边框的左上角相对于文档左上角的位置（以磅为单位）。 |
| Top | float | 必填 | 自选形状边框的左上角相对于文档顶部的位置（以磅为单位）。 |
| Width | float | 必填 | 自选形状边框的宽度（以磅为单位）。 |
| Height | float | 必填 | 自选形状边框的高度（以磅为单位）。 |

## 返回值

Shape。

## 示例

```javascript
/*本示例向第一张工作表中添加矩形。*/
function test() {
    Application.Worksheets.Item(1).Shapes.AddShape(msoShapeRectangle, 50, 50, 100, 200)
}
```

```javascript
/*本示例将十字添加到活动工作表中，然后将其边框设置为8磅厚和红色。*/
function test() {
    let shapes = ActiveSheet.Shapes
    let shape = shapes.AddShape(msoShapeCross, 10, 10, 50, 70)
    shape.Line.Weight = 8
    shape.Line.ForeColor.RGB = RGB(255, 0, 0)
}
```
