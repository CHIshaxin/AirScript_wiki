---
title: "MergeCells 属性"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/DisplayFormat/%E5%B1%9E%E6%80%A7/MergeCells%20%E5%B1%9E%E6%80%A7.html
---
# MergeCells (属性)

返回一个值，该值指示当前用户界面中显示的关联 **[Range](./../../Range/Range 对象.html)** 对象是否包含合并单元格。只读。

## 说明

如果该区域包含合并单元格，则返回 **True**。

## 返回值

Variant

## 示例

```javascript
/*本示例显示 D5:E5 单元格是否包含合并单元格。*/
function test() {
    console.log(Range("D5:E5").DisplayFormat.MergeCells)
}
```

```javascript
/*本示例判断 D7:D10 单元格是否包含合并单元格，并通知用户。*/
function test() {
    if (Range("D7:D10").DisplayFormat.MergeCells) {
        console.log("包含合并单元格")
    } else {
        console.log("未包含合并单元格")
    }
}
```
