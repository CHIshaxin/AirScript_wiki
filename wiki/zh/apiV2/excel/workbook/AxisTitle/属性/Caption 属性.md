---
title: "Caption (属性)"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/AxisTitle/%E5%B1%9E%E6%80%A7/Caption%20%E5%B1%9E%E6%80%A7.html
updated: 1789700916000
hash: uJujxJJo
---

# Caption (属性)

返回或设置一个 **String** 值，它代表坐标轴标题文本。

## 示例

```javascript
/*本示例将工作表 Sheet1 中第一张图表的数值轴标题设置为“一年级成绩”。*/
function test() {
    let axes = Application.Worksheets.Item("Sheet1").ChartObjects(1).Chart.Axes(xlValue)
    axes.HasTitle = true
    let axistitle = axes.AxisTitle
    axistitle.Caption = "一年级成绩"
}
```

```javascript
/*本示例显示图表工作表 Chart1 中图表的分类轴标题。*/
function test() {
    let axistitle = Application.Charts.Item("Chart1").ChartObjects(1).Chart.Axes(xlCategory).AxisTitle
    console.log(axistitle.Caption)
}
```
