---
title: "AxisTitle (对象)"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/AxisTitle/AxisTitle%20%E5%AF%B9%E8%B1%A1.html
updated: 1789700916000
hash: 1iSznHL8
---

# AxisTitle (对象)

代表图表坐标轴标题。

## 说明

使用 **AxisTitle** 属性可返回 **[AxisTitle](./../AxisTitle/AxisTitle 对象.html)** 对象。

只有当坐标轴的 **HasTitle** 属性为 **True** 时， **[AxisTitle](./../AxisTitle/AxisTitle 对象.html)** 对象才存在，从而才能使用该对象。

## 示例

```javascript
/*下例激活第一个嵌入式图表，设置其数值轴标题文本，将其字体设为 10 磅的“Bookman”，并将单词“millions”设为倾斜。*/
function test() {
    Application.Worksheets.Item("Sheet1").ChartObjects(1).Activate()
    let axes = Application.ActiveChart.Axes(xlValue)
    axes.HasTitle = true
    let axistitle = axes.AxisTitle
    axistitle.Caption = "Revenue (millions)"
    axistitle.Font.Name = "bookman"
    axistitle.Font.Size = 10
    axistitle.Characters(10, 8).Font.Italic = true
}
```

```javascript
/*本示例将图表工作表 Chart1 中图表的分类轴标题设置为“考核分类”，并将该标题设置为加粗。*/
function test() {
    let axistitle = Application.Charts.Item("Chart1").ChartObjects(1).Chart.Axes(xlCategory).AxisTitle
    axistitle.Caption = "考核分类"
    axistitle.Characters().Font.Bold = true
}
```
