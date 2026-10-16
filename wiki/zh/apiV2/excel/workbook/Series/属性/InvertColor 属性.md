---
title: "InvertColor (属性)"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/Series/%E5%B1%9E%E6%80%A7/InvertColor%20%E5%B1%9E%E6%80%A7.html
updated: 1789700916000
hash: Dnb_rEaX
---

# InvertColor (属性)

返回或设置系列中负数据点的填充色。可读/写

## 说明

**InvertColor**属性可用来设置负数据点的填充色，颜色可设置为特定的数值、十六进制值、八进制值或 RGB 颜色值。若要将该值设置为 RGB 值，请使用 JavaScript **RGB** 函数。如果不使用 **InvertColor** 属性，还可以使用 **InvertColorIndex**属性，它使用当前调色板上更简单的整数值集合。

为了使 **InvertColor** 属性生效，还必须将 **[Series](./../../Series/Series 对象.html)** 对象的 **InvertIfNegative** 属性设置为 **True**。

## 示例

```javascript
/*下面的代码示例将“Chart 2”第一个系列中的负数据点的填充色设置为洋红。*/
function test() {
    Application.ActiveSheet.ChartObjects("Chart 2").Activate()
    Application.ActiveChart.SeriesCollection(1).InvertIfNegative = true
    Application.ActiveChart.SeriesCollection(1).InvertColor = RGB(255, 0, 255)
}
```

```javascript
/*本示例显示 Chart1 中第二个系列中负数据点的填充颜色是否为红色。*/
function test() {
    let series2 = Application.Charts.Item("Chart1").ChartObjects(1).Chart.SeriesCollection(2)
    console.log(series2.InvertColor == RGB(255, 0, 0))
}
```
