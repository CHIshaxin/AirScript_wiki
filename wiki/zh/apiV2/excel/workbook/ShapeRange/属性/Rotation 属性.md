---
title: "Rotation (属性)"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/ShapeRange/%E5%B1%9E%E6%80%A7/Rotation%20%E5%B1%9E%E6%80%A7.html
updated: 1789700916000
hash: DDOZklFF
---

# Rotation (属性)

返回或设形状的旋转角度（以度为单位）。**Single** 型，可读写。

## 说明

旋转量始终四舍五入到最接近的整数。

## 示例

```javascript
/*本示例将第一张工作表中新建的ShapeRange对象的旋转角度设置为 30 度。*/
function test() {
    let shpRange = Application.Worksheets.Item(1).Shapes.Range([1, 2])
    shpRange.Rotation = 30
}
```

```javascript
/*本示例将活动工作表中新建的ShapeRange对象逆时针旋转 10 度。*/
function test() {
    let shpRange = ActiveSheet.Shapes.Range([2, 3])
    shpRange.Rotation -= 10
}
```
