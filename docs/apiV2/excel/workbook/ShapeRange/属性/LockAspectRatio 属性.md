---
title: "LockAspectRatio 属性"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/ShapeRange/%E5%B1%9E%E6%80%A7/LockAspectRatio%20%E5%B1%9E%E6%80%A7.html
---
# LockAspectRatio (属性)

如果指定的形状在调整大小时其原始比例保持不变，则此属性为 **True**。如果调整大小时可以分别更改形状的高度和宽度，则此属性为 **False**。**MsoTriState** 类型，可读写。

## 说明

| **MsoTriState** 可为以下 **MsoTriState** 常量之一: |
| --- |
| **msoCTrue** |
| **msoFalse**。调整大小时可以分别更改形状的高度和宽度。 |
| **msoTriStateMixed** |
| **msoTriStateToggle** |
| **msoTrue**。指定的形状在调整大小时其原始比例保持不变。 |

## 示例

```javascript
/*本示例向第一张工作表添加一个立方体，用新建的shpRange对象设置在调整大小时其原始比例保持不变。*/
function test() {
    let shapes = Application.Worksheets.Item(1).Shapes
    let shape = shapes.AddShape(msoShapeCube, 50, 50, 100, 200)
    let shpRange = shapes.Range([shape.Name])
    shpRange.LockAspectRatio = msoTrue
}
```

```javascript
/*本示例向活动工作表添加一个矩形，通过shpRange对象设置为调整大小时可以分别更改形状的高度和宽度。*/
function test() {
    let shapes = ActiveSheet.Shapes
    let shape = shapes.AddShape(msoShapeRectangle, 100, 100, 100, 100)
    let shpRange = shapes.Range([shape.Name])
    shpRange.LockAspectRatio = msoFalse
}
```
