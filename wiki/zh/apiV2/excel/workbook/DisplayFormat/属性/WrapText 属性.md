---
title: "WrapText (属性)"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/DisplayFormat/%E5%B1%9E%E6%80%A7/WrapText%20%E5%B1%9E%E6%80%A7.html
updated: 1789700916000
hash: q3eau-cX
---

# WrapText (属性)

返回一个值，该值指示 ET 是否会使当前用户界面中显示的关联 **[Range](./../../Range/Range 对象.html)** 对象的文本自动换行。只读。

## 说明

如果该区域中的所有单元格都自动换行，则返回 **True**。

## 返回值

Variant

## 示例

```javascript
/*本示例显示在当前用户界面中 B3 单元格的文本是否自动换行。*/
function test() {
    console.log(Range("B3").DisplayFormat.WrapText)
}
```

```javascript
/*本示例判断在当前用户界面中 E5:E10 单元格的文本是否自动换行，并通知用户。*/
function test() {
    if (Range("E5:E10").DisplayFormat.WrapText) {
        console.log("文本会自动换行")
    } else {
        console.log("文本不会自动换行")
    }
}
```
