---
title: "Shadow 属性"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/Shape/%E5%B1%9E%E6%80%A7/Shadow%20%E5%B1%9E%E6%80%A7.html
---
# Shadow (属性)

返回一个只读的 **[ShadowFormat](./../../ShadowFormat/ShadowFormat 对象.html)** 对象，它包含指定形状的阴影格式属性。

## 示例

```javascript
/*本示例在第一张工作表中添加一个有阴影的矩形。蓝色浮凸效果的阴影距该矩形的右边 4 磅，下边 3 磅*/
function test() {
    let shapes = Application.Worksheets.Item(1).Shapes
    let shadow = shapes.AddShape(msoShapeRectangle, 10, 10, 150, 90).Shadow
    shadow.Type = msoShadow17
    shadow.ForeColor.RGB = RGB(0, 0, 128)
    shadow.OffsetX = 4
    shadow.OffsetY = 3
}
```

```javascript
/*本示例在活动工作表中添加一个有阴影的六边形。半透明浮凸效果的阴影距该六边形的右边 4 磅，上边 5 磅*/
function test() {
    let shapes = Application.ActiveSheet.Shapes
    let shadow = shapes.AddShape(msoShapeHexagon, 200, 300, 100, 80).Shadow
    shadow.Type = msoShadow17
    shadow.Transparency = 0.5
    shadow.OffsetX = 4
    shadow.OffsetY = -5
}
```
