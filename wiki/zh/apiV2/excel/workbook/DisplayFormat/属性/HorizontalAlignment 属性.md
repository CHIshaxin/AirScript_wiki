---
title: "HorizontalAlignment (属性)"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/DisplayFormat/%E5%B1%9E%E6%80%A7/HorizontalAlignment%20%E5%B1%9E%E6%80%A7.html
updated: 1789700916000
hash: pip-FqdD
---

# HorizontalAlignment (属性)

返回一个值，该值表示当前用户界面中显示的关联 **[Range](./../../Range/Range 对象.html)** 对象采用水平对齐方式。只读。

## 返回值

Variant

## 示例

```javascript
/*本示例显示 E3 单元格的水平对齐方式是否为 xlHAlignGeneral 。*/
function test() {
    console.log(Range("E3").DisplayFormat.HorizontalAlignment == xlHAlignGeneral)
}
```

```javascript
/*本示例判断 E3:E6 单元格的水平对齐方式是否为 xlHAlignLeft ，并通知用户。*/
function test() {
    if (Range("E3:E6").DisplayFormat.HorizontalAlignment == xlHAlignLeft) {
        console.log("水平对齐方式是xlHAlignLeft")
    } else {
        console.log("水平对齐方式不是xlHAlignLeft")
    }
}
```
