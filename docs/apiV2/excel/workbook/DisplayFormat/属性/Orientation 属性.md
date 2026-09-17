---
title: "Orientation 属性"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/DisplayFormat/%E5%B1%9E%E6%80%A7/Orientation%20%E5%B1%9E%E6%80%A7.html
---
# Orientation (属性)

返回一个值，该值表示当前用户界面中显示的关联 **[Range](./../../Range/Range 对象.html)** 对象的文本方向。只读。

## 返回值

Variant

## 示例

```javascript
/*本示例显示 B3 单元格文本方向是否为水平排列。*/
function test() {
    console.log(Range("B3").DisplayFormat.Orientation == xlHorizontal)
}
```

```javascript
/*本示例判断 B5:B10 单元格文本方向是否为向下排列，并通知用户。*/
function test() {
    if (Range("B5:B10").DisplayFormat.Orientation == xlDownward) {
        console.log("文本方向是向下排列")
    } else {
        console.log("文本方向不是向下排列")
    }
}
```
