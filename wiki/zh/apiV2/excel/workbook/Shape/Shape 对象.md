---
title: "Shape (对象)"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/Shape/Shape%20%E5%AF%B9%E8%B1%A1.html
updated: 1789700916000
hash: CVHTaYGZ
---

# Shape (对象)

代表绘图层中的对象，例如自选图形、任意多边形、OLE 对象或图片。

## 说明

**[Shape](./../Shape/Shape 对象.html)** 对象是 **[Shapes](./../Shapes/Shapes 对象.html)** 集合的成员。 **[Shapes](./../Shapes/Shapes 对象.html)** 集合包含某个工作簿中的所有形状。

**注释：**有三个代表形状的对象： **[Shapes](./../Shapes/Shapes 对象.html)** 集合，它代表工作簿中所有的形状； **[ShapeRange](./../ShapeRange/ShapeRange 对象.html)** 集合，它代表工作簿中形状的指定子集（例如， **[ShapeRange](./../ShapeRange/ShapeRange 对象.html)** 对象可以代表工作簿中的形状一和形状四，或者，可以代表工作簿中所有选定的形状）； **[Shape](./../Shape/Shape 对象.html)** 对象，它代表文档中的某一个形状。如果您需要同时处理几个形状，或处理选定区域中的多个形状，请使用 **[ShapeRange](./../ShapeRange/ShapeRange 对象.html)** 集合。

以下各节说明了如何使用 Shape 对象：

- 返回与连接符的端点相连的形状。
- 返回新建的任意多边形。
- 返回组中的单个形状。
- 返回新组成的形状组。
- 返回现有的形状。
- 返回选定区域中的形状。

1.返回与连接符的端点相连的形状

要返回一个代表连接符所连接形状之一的 **[Shape](./../Shape/Shape 对象.html)** 对象，请使用 **BeginConnectedShape** 或 **EndConnectedShape** 属性。

2. 返回新建的任意多边形

使用 **BuildFreeform** 和 **AddNodes** 方法可定义一个新任意多边形的几何特性，使用 **ConvertToShape** 方法可创建任意多边形并返回代表它的 **[Shape](./../Shape/Shape 对象.html)** 对象。

**3.** 返回组中的单个形状

使用 **GroupItems**(*index* )（其中 *index* 是形状的名称或组中的索引号）可返回一个代表一组形状中某一形状的 **[Shape](./../Shape/Shape 对象.html)** 对象。

**4.** 返回新组成的形状组

使用 **Group** 或 **Regroup** 方法可将一系列形状分成一组并返回一个 **[Shape](./../Shape/Shape 对象.html)** 对象，该对象代表新形成的组。在形成了一个组之后，您可以按您处理其他任何形状的方法来处理该组

5. 返回现有的形状

使用 **Shapes**(*index* )（其中 *index* 是形状名称或索引号）可返回代表某个形状的 **[Shape](./../Shape/Shape 对象.html)** 对象。

6. 返回选定区域中的形状

使用 `Selection.ShapeRange`(*index*)（其中 *index* 是形状名称或索引号）可返回一个代表选定区域中的形状的 **[Shape](./../Shape/Shape 对象.html)** 对象。

## 示例

```javascript
/*本示例将第一张工作表中形状一和名为“Rectangle 1”的形状进行水平翻转。*/
function test() {
    let shapes = Application.Worksheets.Item(1).Shapes
    shapes.Item(1).Flip(msoFlipHorizontal)
    shapes.Item("Rectangle 1").Flip(msoFlipHorizontal)
}
```

每个添加到 Shapes 集合的形状将被分配一个默认名称。 若要赋予该形状一个更有意义的名称，请使用 Name 属性。

```javascript
/*本示例向活动工作表添加一个正方形，为其命名为“Red Square”，然后设置其前景色和线条样式。*/
function test() {
    let shapes = ActiveSheet.Shapes
    let shape = shapes.AddShape(msoShapeRectangle, 144, 144, 72, 72)
    shape.Name = "Red Square"
    shape.Fill.ForeColor.RGB = RGB(255, 0, 0)
    shape.Line.DashStyle = msoLineDashDot
}
```
