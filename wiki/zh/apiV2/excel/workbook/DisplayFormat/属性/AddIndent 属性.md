---
title: "AddIndent (属性)"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/DisplayFormat/%E5%B1%9E%E6%80%A7/AddIndent%20%E5%B1%9E%E6%80%A7.html
updated: 1789700916000
hash: bUSn1iBV
---

# AddIndent (属性)

返回一个值，该值指示 ET 是否会在单元格中的文本对齐方式设置为等距分布（水平或垂直）时自动缩进当前用户界面中显示的关联 **[Range](./../../Range/Range 对象.html)** 对象的文本。只读。

## 说明

当单元格中的文本对齐方式已设置为等距分布（水平或垂直）时，如果文本会自动缩进，则返回 **True**。

## 返回值

Variant

## 示例

```javascript
/*此示例显示 B3 单元格文本对齐方式设置为等距分布（水平或垂直）时，文本是否会自动缩进。*/
function test() {
    console.log(Range("B3").DisplayFormat.AddIndent)
}
```

```javascript
/*此示例判断 D5:D10 单元格文本对齐方式为等距分布（水平或垂直）时，文本是否会自动缩进，并通知用户。*/
function test() {
    if (Range("D5:D10").DisplayFormat.AddIndent) {
        console.log("文本会自动缩进")
    } else {
        console.log("文本不会自动缩进")
    }
}
```
