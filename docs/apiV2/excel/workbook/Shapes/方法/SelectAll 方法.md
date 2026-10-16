---
title: "SelectAll 方法"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/Shapes/%E6%96%B9%E6%B3%95/SelectAll%20%E6%96%B9%E6%B3%95.html
---
# SelectAll (方法)

选择指定的 **[Shapes](./../../Shapes/Shapes 对象.html)** 集合中的所有形状。

## 示例

```javascript
/*本示例选择第一张工作表中的所有形状，并创建包含所有这些形状的 ShapeRange 集合。*/
function test() {
    let shapes = Application.Worksheets.Item(1).Shapes
    shapes.SelectAll()
    let sr = Application.Selection.ShapeRange
}
```

```javascript
/*本示例将活动工作表的所有形状的边框粗细设置为 8 磅，颜色为红色。*/
function test() {
    let shapes = ActiveSheet.Shapes
    shapes.SelectAll()
    Selection.ShapeRange.Line.Weight = 8
    Selection.ShapeRange.Line.ForeColor.RGB = RGB(255, 0, 0)
}
```
