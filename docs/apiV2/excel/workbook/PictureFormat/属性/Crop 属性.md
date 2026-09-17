---
title: "Crop 属性"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/PictureFormat/%E5%B1%9E%E6%80%A7/Crop%20%E5%B1%9E%E6%80%A7.html
---
# Crop (属性)

返回一个 **Crop**对象，该对象代表指定的 **[PictureFormat](./../../PictureFormat/PictureFormat 对象.html)** 对象的裁剪设置。

## 示例

```javascript
/*本示例显示活动工作表的第二个图片要裁剪的高度。*/
function test() {
    let shape = ActiveSheet.Shapes.Item(2)
    console.log(shape.PictureFormat.Crop.PictureHeight)
}
```

```javascript
/*本示例将第一张工作表的第三个图片要裁剪的宽度设置为300磅。*/
function test() {
    let shape = Worksheets.Item(1).Shapes.Item(3)
    shape.PictureFormat.Crop.PictureWidth = 300
}
```
