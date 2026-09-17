---
title: "Interior 属性"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/DisplayFormat/%E5%B1%9E%E6%80%A7/Interior%20%E5%B1%9E%E6%80%A7.html
---
# Interior (属性)

返回一个 **[Interior](./../../Interior/Interior 对象.html)** 对象，该对象代表当前用户界面中显示的关联 **[Range](./../../Range/Range 对象.html)** 对象的内部。只读。

## 返回值

Interior

## 示例

```javascript
/*本示例显示 G3 单元格的内部颜色是否为红色。*/
function test() {
    console.log(Range("G3").DisplayFormat.Interior.Color == RGB(255, 0, 0))
}
```

```javascript
/*本示例判断 E1:E5 单元格的内部颜色是否为 xlColorIndexNone ，并通知用户。*/
function test() {
    if (Range("E1:E5").DisplayFormat.Interior.ColorIndex == xlColorIndexNone) {
        console.log("内部颜色是xlColorIndexNone")
    } else {
        console.log("内部颜色不是xlColorIndexNone")
    }
}
```
