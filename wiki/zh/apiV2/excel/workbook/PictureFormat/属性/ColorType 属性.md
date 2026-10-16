---
title: "ColorType (属性)"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/PictureFormat/%E5%B1%9E%E6%80%A7/ColorType%20%E5%B1%9E%E6%80%A7.html
updated: 1789700916000
hash: Bjb_KaYc
---

# ColorType (属性)

返回或设置应用于指定图片或 OLE 对象的颜色转换类型。**MsoPictureColorType** 类型，可读写。

## 示例

```javascript
/*本示例通过 MsoPictureColorType 将第一张工作表上的第一个形状的颜色变换类型设置为灰度。第一个形状必须是图片或 OLE 对象。*/
function test() {
    let worksheet = Worksheets.Item(1)
    worksheet.Shapes.Item(1).PictureFormat.ColorType = msoPictureGrayscale
}
```

```javascript
/*本示例显示活动工作表的第二个形状图片的颜色转换类型是否为默认颜色转换。*/
function test() {
    let shape = ActiveSheet.Shapes.Item(2)
    console.log(shape.PictureFormat.ColorType == msoPictureAutomatic)
}
```
