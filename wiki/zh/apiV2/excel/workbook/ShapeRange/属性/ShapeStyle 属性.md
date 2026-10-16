---
title: "ShapeStyle (属性)"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/ShapeRange/%E5%B1%9E%E6%80%A7/ShapeStyle%20%E5%B1%9E%E6%80%A7.html
updated: 1789700916000
hash: U3O2pXEk
---

# ShapeStyle (属性)

返回或设置 **MsoShapeStyleIndex**类型的值，该值代表形状区域的形状样式。可读/写。

## 示例

```javascript
/*本示例先在第一张工作表创建shpRange对象，然后重新设置对象的形状样式。*/
function test() {
    let shpRange = Application.Worksheets.Item(1).Shapes.Range([1, 2])
    shpRange.ShapeStyle = msoShapeStylePreset10
}
```

```javascript
/*本示例将活动工作表的所有形状重新设置形状样式。*/
function test() {
    let shapes = ActiveSheet.Shapes
    shapes.SelectAll()
    Selection.ShapeRange.ShapeStyle = msoShapeStylePreset20
}
```
