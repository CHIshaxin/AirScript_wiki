---
title: "Visible 属性"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/ShapeRange/%E5%B1%9E%E6%80%A7/Visible%20%E5%B1%9E%E6%80%A7.html
---
# Visible (属性)

返回或设置一个 **MsoTriState** 值，它确定对象是否可见。可读写。

## 示例

```javascript
/*本示例先在第一张工作表中创建shpRange对象，然后将该对象设置为不可见。*/
function test() {
    let shpRange = Application.Worksheets.Item(1).Shapes.Range([1, 2])
    shpRange.Visible = msoFalse
}
```

```javascript
/*本示例将活动工作表中所有形状设置为可见。*/
function test() {
    let shapes = ActiveSheet.Shapes
    shapes.SelectAll()
    Selection.ShapeRange.Visible = msoTrue
}
```
