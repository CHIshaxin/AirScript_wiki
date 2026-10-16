---
title: "ModifyAppliesToRange 方法"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/Databar/%E6%96%B9%E6%B3%95/ModifyAppliesToRange%20%E6%96%B9%E6%B3%95.html
---
# ModifyAppliesToRange (方法)

设置此格式规则所应用于的单元格区域。

## 说明

该区域必须采用 A1 引用样式，并全部包含在 **[FormatConditions](./../../FormatConditions/FormatConditions 对象.html)** 集合的父工作表内。可包括区域操作符（冒号）、相交区域操作符（空格）或合并区域操作符（逗号）。也可以使用货币符号，但会被忽略。

您也可以在区域的任意部分中使用局部定义名称，但该名称必须使用宏语言。

## 参数

| 属性 | 数据类型 | 必填 | 说明 |
| --- | --- | --- | --- |
| Range | Range | 必填 | 此格式规则将应用于的区域。 |

## 示例

```javascript
/*本示例设置活动工作表上区域 A1:A10 中第一个条件格式（数据条）所应用于的单元格区域，并将该数据条的展示方式设置为xlRTL。*/
function test() {
    let databar = ActiveSheet.Range("A1:A10").FormatConditions.Item(1)
    databar.ModifyAppliesToRange(Range("A1:A8"))
    databar.Direction = xlRTL
}
```

```javascript
/*本示例设置活动工作表上区域 A1:A10 中第一个条件格式（数据条）的颜色，并设置该数据条所应用于的单元格区域。*/
function test() {
    let databar = ActiveSheet.Range("A1:A10").FormatConditions.Item(1)
    databar.BarColor.ColorIndex = 3
    databar.ModifyAppliesToRange(Range("A1:A5"))
}
```
