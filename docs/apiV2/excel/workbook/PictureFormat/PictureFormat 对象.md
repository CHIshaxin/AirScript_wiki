---
title: "PictureFormat 对象"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/PictureFormat/PictureFormat%20%E5%AF%B9%E8%B1%A1.html
---
# PictureFormat (对象)

包含应用于图片和 OLE 对象的属性和方法。

## 说明

**LinkFormat** 对象只包含应用于链接 OLE 对象的属性和方法。**OLEFormat** 对象包含应用于 OLE 对象的属性和方法，无论这些对象是不是链接对象。

## 示例

```javascript
/*下例设置了第一张工作表中第一个形状的亮度、对比度和颜色的变换，而且在该形状的底部裁剪了 18 磅。要使此示例执行，则第一个形状必须是图片或 OLE 对象。*/
function test() {
    let worksheet = Worksheets.Item(1)
    let pictureFormat = worksheet.Shapes.Item(1).PictureFormat
    pictureFormat.Brightness = 0.3
    pictureFormat.Contrast = 0.7
    pictureFormat.ColorType = msoPictureGrayscale
    pictureFormat.CropBottom = 18
}
```

```javascript
/*此示例在活动工作表的第二个形状图片的右侧裁剪 10 磅。*/
function test() {
    let shape = ActiveSheet.Shapes.Item(2)
    shape.PictureFormat.CropRight = 10
}
```
