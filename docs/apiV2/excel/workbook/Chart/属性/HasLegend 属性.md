---
title: "HasLegend 属性"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/Chart/%E5%B1%9E%E6%80%A7/HasLegend%20%E5%B1%9E%E6%80%A7.html
---
# HasLegend (属性)

如果图表有图例，则该属性值为 **True**。**Boolean** 类型，可读写。

## 示例

```javascript
/*本示例显示图表工作表 Chart1 的图例，然后将图例的字体颜色设为蓝色。*/
function test() {
    let chart = Application.Charts.Item("Chart1").ChartObjects(1).Chart
    chart.HasLegend = true
    chart.Legend.Font.ColorIndex = 5
}
```

```javascript
/*本示例显示活动工作表第一个嵌入式图表是否有图例。*/
function test() {
    let chart = Application.ActiveSheet.ChartObjects(1).Chart
    console.log(chart.HasLegend)
}
```
