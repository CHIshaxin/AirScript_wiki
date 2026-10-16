---
title: "CopyPicture (方法)"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/Range/%E6%96%B9%E6%B3%95/CopyPicture%20%E6%96%B9%E6%B3%95.html
updated: 1789700916000
hash: B7gikHxb
---

# CopyPicture (方法)

将所选对象作为图片复制到剪贴板。**Variant**。

## 参数

| 属性 | 数据类型 | 必填 | 说明 |
| --- | --- | --- | --- |
| Appearance | Default | 可选 | 指定图片的复制方式。 |

## 返回值

Variant

## 示例

```javascript
/*本示例将 A1:C1 单元格区域的值设置为 2，将 A1:C1 单元格区域复制为图片粘贴到 C3 单元格中，再以位图格式复制，粘贴到 A10 单元格中*/
function test() {
    Range("A1:C1").Value2 = 2
    Range("A1:C1").CopyPicture()
    Range("C3").PasteSpecial()
    Range("A1:C1").CopyPicture(null, xlBitmap)
    Range("A10").PasteSpecial()
}
```

```javascript
/*本示例将 A1 单元格按其打印效果、以矢量格式复制为图片，并粘贴至 E5 单元格中*/
function test() {
    Range("A1").CopyPicture(xlPrinter, xlPicture)
    Range("E5").PasteSpecial()
}
```
