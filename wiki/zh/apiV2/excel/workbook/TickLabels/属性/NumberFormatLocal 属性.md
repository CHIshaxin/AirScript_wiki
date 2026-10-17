---
title: "NumberFormatLocal (属性)"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/TickLabels/%E5%B1%9E%E6%80%A7/NumberFormatLocal%20%E5%B1%9E%E6%80%A7.html
updated: 1789700916000
hash: BnRc_rP-
---

# NumberFormatLocal (属性)

以采用用户语言字符串的形式返回或设置一个 **Variant** 值，它代表对象的格式代码。

## 说明

**Format** 函数使用的格式代码字符串与 **NumberFormat** 和 **NumberFormatLocal** 属性使用的格式代码字符串不同。

## 示例

```javascript
/*下例采用用户语言显示 Sheet1 上的第一个图表分类轴刻度线标签的数字格式。*/
function test() {
    let tickLabels = Application.Worksheets.Item("Sheet1").ChartObjects(1).Chart.Axes(xlCategory).TickLabels
    console.log(`刻度线标签的格式：${tickLabels.NumberFormatLocal}`)
}
```

```javascript
/*以下示例将 Chart1 上数值轴刻度线标签的数字格式（用户语言）设置为“G/通用格式”。*/
function test() {
    let tickLabels = Application.Charts.Item("Chart1").ChartObjects(1).Chart.Axes(xlValue).TickLabels
    tickLabels.NumberFormatLocal = "G/通用格式"
}
```
