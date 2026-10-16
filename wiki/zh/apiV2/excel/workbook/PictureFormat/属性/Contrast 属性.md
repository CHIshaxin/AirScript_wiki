---
title: "Contrast (属性)"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/PictureFormat/%E5%B1%9E%E6%80%A7/Contrast%20%E5%B1%9E%E6%80%A7.html
updated: 1789700916000
hash: B6vSKRa_
---

# Contrast (属性)

返回或设置指定图片或 OLE 对象的对比度。该属性的值必须是 0.0（最低对比度）到 1.0（最高对比度）的数。**Single** 型，可读写。

## 示例

```javascript
/*此示例设置了第一张工作表中第一个形状的对比度。第一个形状必须是图片或 OLE 对象。*/
function test() {
    let worksheet = Worksheets.Item(1)
    worksheet.Shapes.Item(1).PictureFormat.Contrast = 0.8
}
```

```javascript
/*本示例显示活动工作表的第二个形状图片的对比度。*/
function test() {
    let shape = ActiveSheet.Shapes.Item(2)
    console.log(shape.PictureFormat.Contrast)
}
```
