---
title: "PictureFormat (属性)"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/ShapeRange/%E5%B1%9E%E6%80%A7/PictureFormat%20%E5%B1%9E%E6%80%A7.html
updated: 1789700916000
hash: CHen4q3h
---

# PictureFormat (属性)

返回一个 **[PictureFormat](./../../PictureFormat/PictureFormat 对象.html)** 对象，它包含指定形状的图片格式属性。应用于代表图片或 OLE 对象的 **[ShapeRange](./../../ShapeRange/ShapeRange 对象.html)** 对象。只读。

## 示例

```javascript
/*本示例先在第一张工作表创建shpRange对象，然后设置对象的亮度和对比度。*/
function test() {
    let shpRange = Application.Worksheets.Item(1).Shapes.Range([1])
    shpRange.PictureFormat.Brightness = 0.3
    shpRange.PictureFormat.Contrast = 0.75
}
```

```javascript
/*本示例将活动工作表中新建的shpRange对象的颜色转换类型设置为灰度。*/
function test() {
    let shpRange = ActiveSheet.Shapes.Range([2])
    shpRange.PictureFormat.ColorType = msoPictureGrayscale
}
```
