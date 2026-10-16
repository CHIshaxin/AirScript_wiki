---
title: "ShrinkToFit 属性"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/DisplayFormat/%E5%B1%9E%E6%80%A7/ShrinkToFit%20%E5%B1%9E%E6%80%A7.html
---
# ShrinkToFit (属性)

返回一个值，该值指示 ET 是否会自动缩小文本以适合当前用户界面中显示的关联 **[Range](./../../Range/Range 对象.html)** 对象的可用列宽。只读。

## 说明

如果文本会自动缩小以适合可用列宽，则返回 **True**。

## 返回值

Variant

## 示例

```javascript
/*本示例显示 D3 单元格是否自动收缩文本以适应当前用户界面中显示的可用列宽。*/
function test() {
    console.log(Range("D3").DisplayFormat.ShrinkToFit)
}
```

```javascript
/*本示例判断 E5:E10 单元格是否自动收缩文本以适应当前用户界面中显示的可用列宽，并通知用户。*/
function test() {
    if (Range("E5:E10").DisplayFormat.ShrinkToFit) {
        console.log("文本会自动收缩")
    } else {
        console.log("文本不会自动收缩")
    }
}
```
