---
title: "Format 属性"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/LegendKey/%E5%B1%9E%E6%80%A7/Format%20%E5%B1%9E%E6%80%A7.html
---
# Format (属性)

返回 **[ChartFormat](./../../ChartFormat/ChartFormat 对象.html)** 对象。只读。

## 示例

```javascript
/*下例将 Sheet1 上的第一个图表的第三个图例项的图例标示前景填充色设置为绿色。*/
function test() {
    let legend = Application.Worksheets.Item("Sheet1").ChartObjects(1).Chart.Legend
    legend.LegendEntries(3).LegendKey.Format.Fill.ForeColor.RGB = RGB(0, 255, 0)
}
```

```javascript
/*本示例设置图表工作表 Chart1 上第二个图例项的图例标示边框线条的粗细。*/
function test() {
    let legend = Application.Charts.Item("Chart1").ChartObjects(1).Chart.Legend
    legend.LegendEntries(2).LegendKey.Format.Line.Weight = 2
}
```
