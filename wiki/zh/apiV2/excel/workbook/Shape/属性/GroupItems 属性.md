---
title: "GroupItems (属性)"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/Shape/%E5%B1%9E%E6%80%A7/GroupItems%20%E5%B1%9E%E6%80%A7.html
updated: 1789700916000
hash: CQOUqEue
---

# GroupItems (属性)

返回一个 **[GroupShapes](./../../GroupShapes/GroupShapes 对象.html)** 对象，它代表指定形状组中的单个形状。使用 **[GroupShapes](./../../GroupShapes/GroupShapes 对象.html)** 对象的 **Item** 方法可从形状组中返回单个形状。应用于代表成组形状的 **[Shape](./../../Shape/Shape 对象.html)** 对象。只读。

## 示例

```javascript
/*本示例向活动工作表添加三个三角形，将它们组合，为整个组设置颜色，然后仅更改第二个三角形的颜色。*/
function test() {
    let shapes = ActiveSheet.Shapes
    shapes.AddShape(msoShapeIsoscelesTriangle, 10, 10, 100, 100).Name = "shpOne"
    shapes.AddShape(msoShapeIsoscelesTriangle, 150, 10, 100, 100).Name = "shpTwo"
    shapes.AddShape(msoShapeIsoscelesTriangle, 300, 10, 100, 100).Name = "shpThree"
    let groupShape = shapes.Range(["shpOne", "shpTwo", "shpThree"]).Group()
    groupShape.Fill.PresetTextured(msoTextureBlueTissuePaper)
    groupShape.GroupItems.Item(2).Fill.PresetTextured(msoTextureGreenMarble)
}
```

```javascript
/*本示例向第一张工作表添加矩形和平行四边形，将它们组合，然后仅更改第一个矩形的不透明度为50%。*/
function test() {
    let shapes = Application.Worksheets.Item(1).Shapes
    shapes.AddShape(msoShapeRectangle, 10, 10, 50, 50).Name = "shpOne"
    shapes.AddShape(msoShapeParallelogram, 150, 10, 50, 50).Name = "shpTwo"
    let groupShape = shapes.Range(["shpOne", "shpTwo"]).Group()
    groupShape.GroupItems.Item(1).Fill.Transparency = 0.5
}
```
