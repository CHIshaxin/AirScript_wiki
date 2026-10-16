---
title: "Legend 属性"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/Chart/%E5%B1%9E%E6%80%A7/Legend%20%E5%B1%9E%E6%80%A7.html
---
# Legend (属性)

返回一个 **[Legend](./../../Legend/Legend 对象.html)** 对象，该对象表示指定图表的图例。只读。

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
/*本示例判断如果第一张工作表第一个嵌入式图表的图例可见，则将图例字体大小设为12。*/
function test() {
    let chart = Application.Sheets.Item(1).ChartObjects(1).Chart
    if (chart.HasLegend) {
        chart.Legend.Font.Size = 12
    }
}
```
