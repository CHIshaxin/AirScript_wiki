---
title: "Degree (属性)"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/LinearGradient/%E5%B1%9E%E6%80%A7/Degree%20%E5%B1%9E%E6%80%A7.html
updated: 1789700916000
hash: BqiDdaHJ
---

# Degree (属性)

选定区域中线性渐变填充的角度。可读/写。

## 说明

使用 0 - 360 之间的值。

## 返回值

Double

## 示例

```javascript
/*下例将选定区域中线性渐变的填充角度设置为45。*/
function test() {
    Selection.Interior.Pattern = xlPatternLinearGradient
    Selection.Interior.Gradient.Degree = 45
}
```

```javascript
/*下例显示第一张工作表上区域 B1:B10 线性渐变的填充角度。*/
function test() {
    let linearGradient = Worksheets.Item(1).Range("B1:B10").Interior.Gradient
    console.log(linearGradient.Degree)
}
```
