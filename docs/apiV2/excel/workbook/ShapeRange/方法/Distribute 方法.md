---
title: "Distribute 方法"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/ShapeRange/%E6%96%B9%E6%B3%95/Distribute%20%E6%96%B9%E6%B3%95.html
---
# Distribute (方法)

水平或垂直地分布指定的形状区域中的各形状。

## 参数

| 属性 | 数据类型 | 必填 | 说明 |
| --- | --- | --- | --- |
| DistributeCmd | Kso.MsoDistributeCmd | 必填 | 指定该范围内的形状是水平分布还是垂直分布。 |
| RelativeTo | Kso.MsoTriState | 必填 | 不在 ET 中使用。必须为 False。 |

## 示例

```javascript
/*本示例在第一张工作表上定义了一个包含所有自选形状对象的形状子集，然后水平地分布该子集中的形状。最左边的形状将保留在原位。*/
function test() {
    let shapes = Application.Worksheets.Item(1).Shapes
    let autoShpArray = []
    let numShapes = shapes.Count
    if (numShapes > 1) {
        let numAutoShapes = 0
        for (i = 1; i <= numShapes; i++) {
            if (shapes.Item(i).Type == msoAutoShape) {
                numAutoShapes++
                autoShpArray[numAutoShapes] = shapes.Item(i).Name
            }
        }
        if (numAutoShapes > 1) {
            let asRange = shapes.Range(autoShpArray)
            asRange.Distribute(msoDistributeHorizontally, false)
        }
    }
}
```

```javascript
/*本示例在活动工作表中添加三个矩形，然后垂直地分布形状。*/
function test() {
    let shapes = ActiveSheet.Shapes
    let shape1 = shapes.AddShape(msoShapeRectangle, 100, 100, 100, 100)
    let shape2 = shapes.AddShape(msoShapeRectangle, 100, 150, 100, 100)
    let shape3 = shapes.AddShape(msoShapeRectangle, 100, 500, 100, 100)
    let shapeRange = shapes.Range([shape1.Name, shape2.Name, shape3.Name])
    shapeRange.Distribute(msoDistributeVertically, false)
}
```
