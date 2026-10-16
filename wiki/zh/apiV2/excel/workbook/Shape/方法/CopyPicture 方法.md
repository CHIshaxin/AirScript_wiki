---
title: "CopyPicture (方法)"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/Shape/%E6%96%B9%E6%B3%95/CopyPicture%20%E6%96%B9%E6%B3%95.html
updated: 1789700916000
hash: BR1o1EIo
---

# CopyPicture (方法)

将所选对象作为图片复制到剪贴板。

## 说明

如果复制区域，它必须由相邻单元格构成。

## 参数

| 属性 | 数据类型 | 必填 | 说明 |
| --- | --- | --- | --- |
| Appearance | any | 可选 | 用于指定图片复制方式的 XlPictureAppearance 常量。默认值为 **xlScreen**。 |
| Format | any | 可选 | 用于指定图片格式的 XlCopyPictureFormat 常量。默认值为 **xlPicture**。 |

## 示例

```javascript
/*本示例将第一张工作表中的第一个形状作为图片复制到剪贴板。*/
function test() {
    let shapes = Application.Worksheets.Item(1).Shapes
    shapes.Item(1).CopyPicture(xlScreen, xlBitmap)
}
```

```javascript
/*本示例判断如果活动工作表中第二个形状为椭圆，则作为图片复制到剪贴板。*/
function test() {
    let shapes = ActiveSheet.Shapes
    if (shapes.Item(2).AutoShapeType == msoShapeOval) {
        shapes.Item(2).CopyPicture()
    }
}
```
