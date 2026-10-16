---
title: "ShapeRange (对象)"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/ShapeRange/ShapeRange%20%E5%AF%B9%E8%B1%A1.html
updated: 1789700916000
hash: B0onFID-
---

# ShapeRange (对象)

代表形状区域，它是文档中的一组形状。

## 说明

形状区域可以只包含文档中的一个形状，或者也可包含所有形状。您可以在形状区域中包含所需的任意形状（在文档中的所有形状中选取，或从选定内容中的所有形状中选取）。例如，您可以构造一个 **[ShapeRange](./../ShapeRange/ShapeRange 对象.html)** 集合，它包含文档中前三个形状、文档中所有选定的形状，或文档中所有的任意多边形的。

### 1. 返回指定名称或索引号的一组形状

使用 `Shapes.Range`(*index*)（其中 *index* 是形状的名称或索引号，或由形状的名称或索引号组成的数组）可返回代表文档中的一组形状的 **[ShapeRange](./../ShapeRange/ShapeRange 对象.html)** 集合。您可以使用 **Array** 函数来构造名称或索引号的数组。

## 示例

```javascript
/*本示例设置第一张工作表上的形状一和三的填充图案。*/
function test() {
    let myDocument = Application.Worksheets.Item(1)
    myDocument.Shapes.Range([1, 3]).Fill.Patterned(msoPatternHorizontalBrick)
}
```

虽然可以使用 **Range** 属性来返回任意数量的形状或幻灯片，但如果您只想返回一个集合成员，则使用 **Item** 方法会更简单。例如，`Shapes(1)` 比 `Shapes.Range(1)` 简单。

```javascript
/*本示例设置第一张工作表上名为“Oval 4”和“Rectangle 5”的形状的填充图案。*/
function test() {
    let myDocument = Application.Worksheets.Item(1)
    let myRange = myDocument.Shapes.Range(["Oval 4", "Rectangle 5"])
    myRange.Fill.Patterned(msoPatternHorizontalBrick)
}
```

### 2. 返回文档中全部或部分选定的形状

使用 **Selection** 对象的 **ShapeRange** 属性可返回选定对象中的所有形状。

```javascript
/*本示例为活动工作表上选定内容中的所有形状设置前景填充色。*/
function test() {
    let shapes = ActiveSheet.Shapes
    shapes.SelectAll()
    Selection.ShapeRange.Fill.ForeColor.RGB = RGB(255, 0, 255)
}
```

使用 `Selection.ShapeRange`(*index*)（其中 *index* 是形状的名称或索引号）返回某一选定的形状。

```javascript
/*本示例为活动工作表上选定内容中的第二个形状设置前景填充色。*/
function test() {
    let shapes = ActiveSheet.Shapes
    shapes.SelectAll()
    Selection.ShapeRange.Item(2).Fill.ForeColor.RGB = RGB(255, 0, 255)
}
```
