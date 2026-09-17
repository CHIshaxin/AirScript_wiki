---
title: "NumberFormat 属性"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/DisplayFormat/%E5%B1%9E%E6%80%A7/NumberFormat%20%E5%B1%9E%E6%80%A7.html
---
# NumberFormat (属性)

返回一个值，该值表示当前用户界面中显示的关联 **[Range](./../../Range/Range 对象.html)** 对象的格式代码。只读。

## 说明

如果指定区域中的所有单元格不具有相同的数字格式，则此属性将返回 **Null**。

## 返回值

Variant

## 示例

```javascript
/*本示例显示 F5 单元格的数字格式。*/
function test() {
    console.log(Range("F5").DisplayFormat.NumberFormat)
}
```

```javascript
/*本示例判断 D5:D10 单元格的数字格式是否为 General ，并通知用户。*/
function test() {
    if (Range("D5:D10").DisplayFormat.NumberFormat == "General") {
        console.log("数字格式为General")
    } else {
        console.log("数字格式不是General")
    }
}
```
