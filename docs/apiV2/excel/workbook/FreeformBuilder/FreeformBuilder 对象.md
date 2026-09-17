---
title: "FreeformBuilder 对象"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/FreeformBuilder/FreeformBuilder%20%E5%AF%B9%E8%B1%A1.html
---
# FreeformBuilder (对象)

代表任意多边形创建时的几何属性。

## 说明

使用 **BuildFreeform** 方法可返回 **[FreeformBuilder](./../FreeformBuilder/FreeformBuilder 对象.html)** 对象。使用 **AddNodes** 方法可在任意多边形中添加节点。使用 **ConvertToShape** 方法可创建 **[FreeformBuilder](./../FreeformBuilder/FreeformBuilder 对象.html)** 对象中定义的形状，并将它添加到 **[Shapes](./../Shapes/Shapes 对象.html)** 集合中。

## 示例

```javascript
/*本示例向第一张工作表中添加带有四条线段的任意多边形。*/
function test() {
    let builder = Worksheets.Item(1).Shapes.BuildFreeform(msoEditingCorner, 360, 200)
    builder.AddNodes(msoSegmentCurve, msoEditingCorner, 380, 230, 400, 250, 450, 300)
    builder.AddNodes(msoSegmentCurve, msoEditingAuto, 480, 200)
    builder.AddNodes(msoSegmentLine, msoEditingAuto, 480, 400)
    builder.AddNodes(msoSegmentLine, msoEditingAuto, 360, 200)
    let s = builder.ConvertToShape()
    s.Fill.ForeColor.RGB = RGB(0, 0, 0)
}
```

```javascript
/*本示例在活动工作表中创建一个具有三个顶点的红色任意多边形。*/
function test() {
    let builder = ActiveSheet.Shapes.BuildFreeform(msoEditingAuto, 470, 330)
    builder.AddNodes(msoSegmentCurve, msoEditingSmooth, 570, 360)
    builder.AddNodes(msoSegmentCurve, msoEditingSmooth, 690, 430)
    builder.AddNodes(msoSegmentCurve, msoEditingSmooth, 470, 330)
    let s = builder.ConvertToShape()
    s.Fill.ForeColor.RGB = RGB(255, 0, 0)
}
```
