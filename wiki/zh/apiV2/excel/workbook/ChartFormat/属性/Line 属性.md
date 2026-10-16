---
title: "Line (属性)"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/ChartFormat/%E5%B1%9E%E6%80%A7/Line%20%E5%B1%9E%E6%80%A7.html
updated: 1789700916000
hash: 0w2sKQAw
---

# Line (属性)

返回一个包含指定图表元素线条格式属性的 **[LineFormat](./../../LineFormat/LineFormat 对象.html)** 对象。只读。

## 说明

对于线条来说， **[LineFormat](./../../LineFormat/LineFormat 对象.html)** 对象代表线条本身；而对于带有边框的图表来说， **[LineFormat](./../../LineFormat/LineFormat 对象.html)** 对象代表边框。

## 示例

```javascript
/*此示例设置工作表 Sheet1 中第一个内嵌图表区的线条粗细为 1.6 磅。*/
function test() {
    let chartformat = Application.Worksheets.Item("Sheet1").ChartObjects(1).Chart.ChartArea.Format
    chartformat.Line.Weight = 1.6
}
```

```javascript
/*此示例显示图表工作表 Chart1 图表区线条的虚线样式是否为实线。*/
function test() {
    let chartformat = Application.Charts.Item("Chart1").ChartObjects(1).Chart.ChartArea.Format
    console.log(chartformat.Line.DashStyle == msoLineSolid)
}
```
