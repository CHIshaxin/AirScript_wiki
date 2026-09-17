---
title: "IncrementTop 方法"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/ShapeRange/%E6%96%B9%E6%B3%95/IncrementTop%20%E6%96%B9%E6%B3%95.html
---
# IncrementTop (方法)

以指定磅数垂直移动指定形状。

## 参数

| 属性 | 数据类型 | 必填 | 说明 |
| --- | --- | --- | --- |
| Increment | float | 必填 | 指定形状对象垂直移动的距离，以磅为单位。正值将形状下移，负值将形状上移。 |

## 示例

```javascript
/*本示例将第一张工作表中新建的ShapeRange对象设置填充，再向右移 70 磅、向上移 50 磅，然后顺时针旋转 30 度。*/
function test() {
    let shpRange = Application.Worksheets.Item(1).Shapes.Range([1, 2])
    shpRange.Fill.PresetTextured(msoTextureGranite)
    shpRange.IncrementLeft(70)
    shpRange.IncrementTop(-50)
    shpRange.IncrementRotation(30)
}
```

```javascript
/*本示例将活动工作表的所有形状向下移 190 磅，然后逆时针旋转 180 度。*/
function test() {
    let shapes = ActiveSheet.Shapes
    shapes.SelectAll()
    Selection.ShapeRange.IncrementTop(190)
    Selection.ShapeRange.IncrementRotation(-180)
}
```
