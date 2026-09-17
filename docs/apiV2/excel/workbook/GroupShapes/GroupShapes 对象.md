---
title: "GroupShapes 对象"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/GroupShapes/GroupShapes%20%E5%AF%B9%E8%B1%A1.html
---
# GroupShapes (对象)

代表一组形状中的单个形状。

## 说明

每个形状都由一个 **[Shape](./../Shape/Shape 对象.html)** 对象代表。将 **Item**方法与此对象一起使用，您可以在不取消分组的情况下处理组合的各个形状。

使用 **GroupItems** 属性可返回 **[GroupShapes](./../GroupShapes/GroupShapes 对象.html)** 集合。使用 **GroupItems**(*index*)（其中 *index* 是分组的形状中单个形状的编号）可从 **[GroupShapes](./../GroupShapes/GroupShapes 对象.html)** 集合中返回单个形状。

```javascript
/*下例向 myDocument 添加三个三角形，将它们分成一组，设置整个组的颜色，然后只更改第二个三角形的颜色。*/
function test(){
let myDocument = Application.Worksheets.Item(1)
let shape1 = myDocument.Shapes
    shape1.AddShape(msoShapeIsoscelesTriangle, 10, 10, 100, 100).Name = "shpOne"
    shape1.AddShape(msoShapeIsoscelesTriangle, 150, 10, 100, 100).Name = "shpTwo"
    shape1.AddShape(msoShapeIsoscelesTriangle, 300, 10, 100, 100).Name = "shpThree"
    let shape2 = shape1.Range(["shpOne", "shpTwo", "shpThree"]).Group()
        shape2.Fill.PresetTextured(msoTextureBlueTissuePaper)
        shape2.GroupItems.Item(2).Fill.PresetTextured(msoTextureGreenMarble)
}
```
