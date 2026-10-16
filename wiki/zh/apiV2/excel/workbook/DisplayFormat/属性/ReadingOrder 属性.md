---
title: "ReadingOrder (属性)"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/DisplayFormat/%E5%B1%9E%E6%80%A7/ReadingOrder%20%E5%B1%9E%E6%80%A7.html
updated: 1789700916000
hash: L6leW4K0
---

# ReadingOrder (属性)

返回当前用户界面中显示的关联 **[Range](./../../Range/Range 对象.html)** 对象的读取顺序。只读。

## 返回值

Long

## 示例

```javascript
/*本示例显示 F4 单元格的读取顺序是否根据上下文。*/
function test() {
    console.log(Range("F4").DisplayFormat.ReadingOrder == xlContext)
}
```

```javascript
/*本示例判断 D10 单元格的读取顺序是否为从左到右，并通知用户。*/
function test() {
    if (Range("D10").DisplayFormat.ReadingOrder == xlLTR) {
        console.log("读取顺序是从左到右")
    } else {
        console.log("读取顺序不是从左到右")
    }
}
```
