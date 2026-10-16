---
title: "BuildFreeform (方法)"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/Shapes/%E6%96%B9%E6%B3%95/BuildFreeform%20%E6%96%B9%E6%B3%95.html
updated: 1789700916000
hash: DckUsdD_
---

# BuildFreeform (方法)

建立一个任意多边形对象。返回一个 **[FreeformBuilder](./../../FreeformBuilder/FreeformBuilder 对象.html)** 对象，该对象代表正在创建的任意多边形。用 **AddNodes** 方法向任意多边形添加线段。如果任意多边形中已包含了一个以上的线段，则可用 **ConvertToShape** 方法将 **[FreeformBuilder](./../../FreeformBuilder/FreeformBuilder 对象.html)** 对象转换为 **[Shape](./../../Shape/Shape 对象.html)** 对象，该对象将具有在 **[FreeformBuilder](./../../FreeformBuilder/FreeformBuilder 对象.html)** 对象中定义的几何属性。

## 参数

| 属性 | 数据类型 | 必填 | 说明 |
| --- | --- | --- | --- |
| EditingType | Kso.MsoEditingType | 必填 | 第一个节点的编辑属性。 |
| X1 | float | 必填 | 任意图形中第一个节点相对于文档左上角的位置（以磅为单位）。 |
| Y1 | float | 必填 | 任意图形中第一个节点相对于文档左上角的位置（以磅为单位）。 |

## 返回值

FreeformBuilder

## 示例

```javascript
/*本示例将一个有五个顶点的黑色任意多边形添加到第一张工作表中。*/
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
    let builder = ActiveSheet.Shapes.BuildFreeform(msoEditingSmooth, 380, 220)
    builder.AddNodes(msoSegmentCurve, msoEditingAuto, 460, 260)
    builder.AddNodes(msoSegmentCurve, msoEditingAuto, 480, 320)
    builder.AddNodes(msoSegmentCurve, msoEditingAuto, 380, 220)
    let s = builder.ConvertToShape()
    s.Fill.ForeColor.RGB = RGB(255, 0, 0)
}
```
