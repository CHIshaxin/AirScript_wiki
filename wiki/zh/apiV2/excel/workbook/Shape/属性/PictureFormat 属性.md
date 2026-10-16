---
title: "PictureFormat (属性)"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/Shape/%E5%B1%9E%E6%80%A7/PictureFormat%20%E5%B1%9E%E6%80%A7.html
updated: 1789700916000
hash: CLCWmN_q
---

# PictureFormat (属性)

返回一个 **[PictureFormat](./../../PictureFormat/PictureFormat 对象.html)** 对象，它包含指定形状的图片格式属性。应用于代表图片或 OLE 对象的 **[Shape](./../../Shape/Shape 对象.html)** 对象。只读。

## 示例

```javascript
/*本示例设置了活动工作表中第一个形状的亮度和对比度。第一个形状必须是图片或 OLE 对象。*/
function test() {
    let shape = ActiveSheet.Shapes.Item(1)
    shape.PictureFormat.Brightness = 0.3
    shape.PictureFormat.Contrast = 0.75
}
```

```javascript
/*本示例将第一张工作表中第二个形状颜色转换设置为灰度。第二个形状必须是图片或 OLE 对象。*/
function test() {
    let shape = ActiveSheet.Shapes.Item(2)
    shape.PictureFormat.ColorType = msoPictureGrayscale
}
```
