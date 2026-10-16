---
title: "CropBottom (属性)"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/PictureFormat/%E5%B1%9E%E6%80%A7/CropBottom%20%E5%B1%9E%E6%80%A7.html
updated: 1789700916000
hash: JyI7P1sk
---

# CropBottom (属性)

返回或设置从指定图片或 OLE 对象的底部所裁剪下的磅数。**Single** 型，可读写。

## 说明

裁剪的计算是同图像的原始尺寸相关的。例如，如果插入一幅原始高度为 100 磅的图像，将其放大为 200 磅高，然后设置 **CropBottom** 属性为 50, 那么从图像的底部切除的将是 100 磅（而不是 50 磅）。

## 示例

```javascript
/*此示例在第一张工作表中第三个形状的底部裁剪了 20 磅。要使此示例执行，第三个形状必须是图片或 OLE 对象。*/
function test() {
    let worksheet = Worksheets.Item(1)
    worksheet.Shapes.Item(3).PictureFormat.CropBottom = 20
}
```

```javascript
/*使用此示例可指定从选定的形状底部裁去的百分比，不管是否已调整了该形状的大小。为使此示例正常运行，选定的形状必须是图片或 OLE 对象。*/
function test() {
    let percentToCrop = "What percentage do you want to crop off the bottom of this picture?"
    ActiveSheet.Shapes.SelectAll()
    let shapeToCrop = Selection.ShapeRange.Item(1)
    let duplicate = shapeToCrop.Duplicate()
    duplicate.ScaleHeight(1, true)
    let origHeight = duplicate.Height
    duplicate.Delete()
    let cropPoints = origHeight * percentToCrop / 100
    shapeToCrop.PictureFormat.CropBottom = cropPoints
}
```
