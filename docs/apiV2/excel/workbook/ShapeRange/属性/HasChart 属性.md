---
title: "HasChart 属性"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/ShapeRange/%E5%B1%9E%E6%80%A7/HasChart%20%E5%B1%9E%E6%80%A7.html
---
# HasChart (属性)

返回形状区域是否包含图表。**MsoTriState** 类型，只读。

## 示例

```javascript
/*本示例显示第一张工作表中新建的ShapeRange对象是否包含图表。*/
function test() {
    let shpRange = Application.Worksheets.Item(1).Shapes.Range([2, 3])
    if (shpRange.HasChart) {
        console.log("包含图表")
    }else {
        console.log("不包含图表")
    }
}
```

```javascript
/*本示例显示活动工作表中的所有形状是否包含图表。*/
function test() {
    let shapes = ActiveSheet.Shapes
    shapes.SelectAll()
    if (Selection.ShapeRange.HasChart) {
        console.log("包含图表")
    }else {
        console.log("不包含图表")
    }
}
```
