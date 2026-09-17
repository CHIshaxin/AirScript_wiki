---
title: "PercentMax 属性"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/Databar/%E5%B1%9E%E6%80%A7/PercentMax%20%E5%B1%9E%E6%80%A7.html
---
# PercentMax (属性)

返回或设置一个以单元格宽度百分比形式指定最长数据条长度的 **Long** 值。

## 说明

该值必须是介于 0 和 100 之间的整数。默认值为 100。

**PercentMax** 属性的效果取决于 **[Databar](./../../Databar/Databar 对象.html)** 对象的 **AxisPosition** 属性的设置。当 **AxisPosition** 属性为 **xlDataBarAxisAutomatic**，并且范围同时包含正负值时，最长正数数据条和最长负数数据条的长度总和不会超过 **PercentMax** 属性指定的值。当 **AxisPosition** 属性为 **xlDataBarAxisMidpoint** 时，最长数据条（正或负）将等于 **PercentMax** 属性的值除以 2。当 **AxisPosition** 属性为 **xlDataBarAxisNone**，那么最长数据条的长度总是为 **PercentMax** 属性指定的单元格宽度百分比。

## 示例

```javascript
/*本示例以单元格宽度百分比形式显示活动工作表上区域 A1:A10 中第一个条件格式（数据条）的最长数据条长度。*/
function test() {
    let databar = ActiveSheet.Range("A1:A10").FormatConditions.Item(1)
    console.log(databar.PercentMax)
}
```

```javascript
/*本示例以单元格宽度百分比形式设置活动工作表上区域 A1:A10 中第一个条件格式（数据条）的最长数据条长度。*/
function test() {
    let databar = ActiveSheet.Range("A1:A10").FormatConditions.Item(1)
    databar.PercentMax = 80
}
```
