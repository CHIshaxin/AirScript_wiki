---
title: "ConvertToShape 方法"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/FreeformBuilder/%E6%96%B9%E6%B3%95/ConvertToShape%20%E6%96%B9%E6%B3%95.html
---
# ConvertToShape (方法)

创建一个具有指定 **[FreeformBuilder](./../../FreeformBuilder/FreeformBuilder 对象.html)** 对象的几何特性的形状。返回一个代表新形状的 **[Shape](./../../Shape/Shape 对象.html)** 对象。

## 说明

对于 **[FreeformBuilder](./../../FreeformBuilder/FreeformBuilder 对象.html)** 对象，至少要应用一次 **AddNodes**方法，然后才能使用 **ConvertToShape** 方法。

## 返回值

Shape

## 示例

```javascript
/*本示例将一个有五个顶点的黑色任意多边形添加到第一张工作表中。*/
function test() {
    let myDocument = Application.Worksheets.Item(1)
    let builder = Application.Worksheets.Item(1).Shapes.BuildFreeform(msoEditingCorner, 360, 200)
    builder.AddNodes(msoSegmentCurve, msoEditingCorner, 380, 230, 400, 250, 450, 300)
    builder.AddNodes(msoSegmentCurve, msoEditingAuto, 480, 200)
    builder.AddNodes(msoSegmentLine, msoEditingAuto, 480, 400)
    builder.AddNodes(msoSegmentLine, msoEditingAuto, 360, 200)
    builder.ConvertToShape().Fill.ForeColor.RGB = RGB(0, 0, 0)
}
```

```javascript
/*本示例在活动工作表中创建一个具有四个顶点的绿色任意多边形。*/
function test() {
    let builder = ActiveSheet.Shapes.BuildFreeform(msoEditingSmooth, 375, 235)
    builder.AddNodes(msoSegmentLine, msoEditingAuto, 478, 265)
    builder.AddNodes(msoSegmentLine, msoEditingAuto, 496, 333)
    builder.AddNodes(msoSegmentLine, msoEditingAuto, 599, 373)
    builder.AddNodes(msoSegmentLine, msoEditingAuto, 375, 235)
    let s = builder.ConvertToShape()
    s.Fill.ForeColor.RGB = RGB(0, 255, 0)
}
```
