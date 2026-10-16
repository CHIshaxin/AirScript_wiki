---
title: "Shapes (对象)"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/Shapes/Shapes%20%E5%AF%B9%E8%B1%A1.html
updated: 1789700916000
hash: CiYRJWst
---

# Shapes (对象)

指定的工作表上的所有 **[Shape](./../Shape/Shape 对象.html)** 对象的集合。

## 说明

每个 **[Shape](./../Shape/Shape 对象.html)** 对象都代表绘图层中的一个对象，如自选图形、任意多边形、OLE 对象或图片。

**注释：**如果您想处理文档中的一部分形状（例如，只针对文档中的自选图形或只对选定的形状进行一些操作），则必须构造一个 **[ShapeRange](./../ShapeRange/ShapeRange 对象.html)** 集合，其中包含您要处理的形状。

使用 **Shapes** 属性可返回 **[Shapes](./../Shapes/Shapes 对象.html)** 集合。下例选定 myDocument 上的所有形状。

**注释：**如果您要同时对工作表上的所有形状进行操作（例如删除或设置属性），请选定所有形状，然后对选定区域使用 **ShapeRange** 属性，以创建一个 **[ShapeRange](./../ShapeRange/ShapeRange 对象.html)** 对象，该对象包含工作表上的所有形状，然后对 **[ShapeRange](./../ShapeRange/ShapeRange 对象.html)** 对象应用相应的属性或方法。

## 示例

```javascript
/*本示例选中第一张工作表中所有形状。*/
function test() {
    let shapes = Application.Worksheets.Item(1).Shapes
    shapes.SelectAll()
}
```

**使用 `Shapes`(*index*)（其中 *index* 是形状的名称或索引号）可返回一个 Shape 对象。下例设置 `myDocument` 上形状一的预设阴影的填充。**

```javascript
/*本示例设置第一张工作表上形状一的预设阴影的填充。*/
function test() {
    let shapes = Application.Worksheets.Item(1).Shapes
    shapes.Item(1).Fill.PresetGradient(msoGradientHorizontal, 1, msoGradientBrass)
}
```

**使用 `Shapes.Range`(*index*)（其中 *index* 是形状的名称或索引号，或是它们的一个数组）可返回一个 ShapeRange 集合，该集合代表 Shapes 集合的一个子集。下例设置 *myDocument* 上形状一和三的填充图案。**

```javascript
/*本示例设置第一张工作表上形状一和三的填充图案。*/
function test() {
    let shapes = Application.Worksheets.Item(1).Shapes
    shapes.Range([1, 3]).Fill.Patterned(msoPatternHorizontalBrick)
}
```

```javascript
/*本示例将活动工作表上名为“Check Box 1”的复选框设置为选中状态。*/
function test() {
    Application.ActiveSheet.Shapes("Check Box 1").ControlFormat.Value = true
}
```
