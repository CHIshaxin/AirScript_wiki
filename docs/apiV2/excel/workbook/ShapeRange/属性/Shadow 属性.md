---
title: "Shadow 属性"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/ShapeRange/%E5%B1%9E%E6%80%A7/Shadow%20%E5%B1%9E%E6%80%A7.html
---
# Shadow (属性)

返回一个只读的 **[ShadowFormat](./../../ShadowFormat/ShadowFormat 对象.html)** 对象，它包含指定形状的阴影格式属性。

## 示例

```javascript
/*本示例先在第一张工作表创建shpRange对象，然后添加距该对象右边 4 磅，下边 3 磅的红色浮凸效果的阴影。*/
function test() {
    let shpRange = Application.Worksheets.Item(1).Shapes.Range([1, 2])
    let shadow = shpRange.Shadow
    shadow.Type = msoShadow17
    shadow.ForeColor.RGB = RGB(255, 0, 0)
    shadow.OffsetX = 4
    shadow.OffsetY = 3
}
```

```javascript
/*本示例将活动工作表的所有形状添加一个半透明浮凸效果的阴影距该形状的右边 4 磅，上边 5 磅。*/
function test() {
    let shapes = ActiveSheet.Shapes
    shapes.SelectAll()
    Selection.ShapeRange.Shadow.Type = msoShadow17
    Selection.ShapeRange.Shadow.Transparency = 0.5
    Selection.ShapeRange.Shadow.OffsetX = 4
    Selection.ShapeRange.Shadow.OffsetY = -5
}
```
