---
title: "AddAboveAverage (方法)"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/FormatConditions/%E6%96%B9%E6%B3%95/AddAboveAverage%20%E6%96%B9%E6%B3%95.html
updated: 1789700916000
hash: DSgd6wBG
---

# AddAboveAverage (方法)

返回一个代表指定区域条件格式规则的新 **[AboveAverage](./../../AboveAverage/AboveAverage 对象.html)** 对象。返回一个 **[AboveAverage](./../../AboveAverage/AboveAverage 对象.html)** 对象。

## 说明

**[AboveAverage](./../../AboveAverage/AboveAverage 对象.html)** 对象用于查找单元格区域中高于或低于平均或标准偏差的值。例如，可在年度业绩评估中查找高于平均业绩的人员，或者在质量评估中查找低于两个标准偏差的加工原料。

## 示例

```javascript
/*本示例在活动工作表上单元格区域 C1:C10 上新增AboveAverage条件格式，并设置该条件格式的字体样式。*/
function test() {
    let aboveAverage = ActiveSheet.Range("C1:C10").FormatConditions.AddAboveAverage()
    aboveAverage.Font.Bold = true
    aboveAverage.Font.Italic = true
}
```

```javascript
/*本示例在活动工作表上单元格区域 A1:A10 上新增AboveAverage条件格式，并设置该条件格式的数字格式。*/
function test() {
    let aboveAverage = ActiveSheet.Range("A1:A10").FormatConditions.AddAboveAverage()
    aboveAverage.NumberFormat = "0.00%"
}
```
