---
title: "Top (属性)"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/LegendKey/%E5%B1%9E%E6%80%A7/Top%20%E5%B1%9E%E6%80%A7.html
updated: 1789700916000
hash: DQ5Srgsr
---

# Top (属性)

返回一个 **Double**值，它代表从对象的上边缘到工作表第一行顶部或图表上的图表区顶部的距离（以磅为单位）。

## 示例

```javascript
/*下例显示 Sheet1 中的第一个图表的第二个图例项的图例标示上边缘到到图表区顶部的距离。*/
function test() {
    let legend = Application.Worksheets.Item("Sheet1").ChartObjects(1).Chart.Legend
    console.log(legend.LegendEntries(2).LegendKey.Top)
}
```

```javascript
/*本示例显示 Chart1 上第一个图例项的图例标示上边缘到到图表区顶部的距离。*/
function test() {
    console.log(Application.Charts.Item("Chart1").ChartObjects(1).Chart.Legend.LegendEntries(1).LegendKey.Top)
}
```
