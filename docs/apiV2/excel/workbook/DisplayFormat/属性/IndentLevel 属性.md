---
title: "IndentLevel 属性"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/DisplayFormat/%E5%B1%9E%E6%80%A7/IndentLevel%20%E5%B1%9E%E6%80%A7.html
---
# IndentLevel (属性)

返回一个值，该值表示当前用户界面中显示的关联 **[Range](./../../Range/Range 对象.html)** 对象的缩进级别。只读。

## 返回值

Variant

## 示例

```javascript
/*本示例显示 E3 单元格的缩进级别。*/
function test() {
    console.log(Range("E3").DisplayFormat.IndentLevel)
}
```

```javascript
/*本示例判断 G5:G10 单元格的缩进级别是否为3，并通知用户。*/
function test() {
    if (Range("G5:G10").DisplayFormat.IndentLevel == 3) {
        console.log("缩进级别为3")
    } else {
        console.log("缩进级别不为3")
    }
}
```
