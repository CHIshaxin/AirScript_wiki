---
title: "Glow (属性)"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/ShapeRange/%E5%B1%9E%E6%80%A7/Glow%20%E5%B1%9E%E6%80%A7.html
updated: 1789700916000
hash: ZGNV-FvG
---

# Glow (属性)

返回一个包含形状区域发光格式属性的指定形状区域的 **GlowFormat** 对象。只读。

## 说明

发光效果为图形添加了色彩鲜明的彩色边缘。

## 示例

```javascript
/*本示例为第一张工作表中新建的ShapeRange对象添加一个蓝色的发光效果，透明度为0.5。*/
function test() {
    let shpRange = Application.Worksheets.Item(1).Shapes.Range([1, 2])
    shpRange.Glow.Color.RGB = RGB(0, 0, 255)
    shpRange.Glow.Transparency = 0.5
}
```

```javascript
/*本示例将活动工作表的所有形状添加一个红色的发光效果，大小为10。*/
function test() {
    let shapes = ActiveSheet.Shapes
    shapes.SelectAll()
    Selection.ShapeRange.Glow.Color.RGB = RGB(255, 0, 0)
    Selection.ShapeRange.Glow.Size = 10
}
```
