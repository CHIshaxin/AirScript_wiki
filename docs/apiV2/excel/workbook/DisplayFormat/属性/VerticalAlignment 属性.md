---
title: "VerticalAlignment 属性"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/DisplayFormat/%E5%B1%9E%E6%80%A7/VerticalAlignment%20%E5%B1%9E%E6%80%A7.html
---
# VerticalAlignment (属性)

返回一个值，该值表示当前用户界面中显示的关联 **[Range](./../../Range/Range 对象.html)** 对象采用垂直对齐方式。只读。

## 返回值

Variant

## 示例

```javascript
/*本示例显示在当前用户界面中 F3 单元格的垂直对齐方式是否为居中对齐。*/
function test() {
    console.log(Range("F3").DisplayFormat.VerticalAlignment == xlVAlignCenter)
}
```

```javascript
/*本示例判断在当前用户界面中 D5:D10 区域单元格的垂直对齐方式是否为顶部对齐，并通知用户。*/
function test() {
    if (Range("D5:D10").DisplayFormat.VerticalAlignment == xlVAlignTop) {
        console.log("垂直对齐方式是顶部对齐")
    } else {
        console.log("垂直对齐方式不是顶部对齐")
    }
}
```
