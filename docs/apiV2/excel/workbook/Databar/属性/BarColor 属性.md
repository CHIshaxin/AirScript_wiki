---
title: "BarColor 属性"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/Databar/%E5%B1%9E%E6%80%A7/BarColor%20%E5%B1%9E%E6%80%A7.html
---
# BarColor (属性)

返回一个可用于修改数据条条件格式中的数据条颜色的 **[FormatColor](./../../FormatColor/FormatColor 对象.html)** 对象。

## 示例

```javascript
/*本示例设置活动工作表上区域 A1:A10 中第一个为数据条的条件格式的主题颜色。*/
function test() {
    let databar = ActiveSheet.Range("A1:A10").FormatConditions.Item(1)
    databar.BarColor.ThemeColor = xlThemeColorAccent3
}
```

```javascript
/*本示例设置活动工作表上区域 A1:A10 中第一个为数据条的条件格式的颜色和该数据条边框的颜色。*/
function test() {
    let databar = ActiveSheet.Range("A1:A10").FormatConditions.Item(1)
    databar.BarBorder.Color.ColorIndex = 3
    databar.BarColor.ColorIndex = 3
}
```
