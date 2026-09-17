---
title: "SeriesLines 对象"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/SeriesLines/SeriesLines%20%E5%AF%B9%E8%B1%A1.html
---
# SeriesLines (对象)

代表图表组中的系列线。

## 说明

系列线连接每个系列中的数据。只有二维堆积条形图、二堆, 2-D 堆只柱形图、复合饼图或复合条饼图等图表可以有系列线。此对象不是集合。没有代表单个系列线的对象；您或者打开图表组中所有数据点的系列线，或者将其全部关闭。

如果 **HasSeriesLines**属性是 **False**， **[SeriesLines](./../SeriesLines/SeriesLines 对象.html)** 对象的绝大部分属性都会被禁用。

## 示例

```javascript
/*本示例将系列线添加到第一张工作表第一个嵌入图表中的第一个图表组中，并将颜色设置为蓝色，该图表必须是二维堆积条形图或柱形图。*/
function test() {
    let chartgroup = Worksheets(1).ChartObjects(1).Chart.ChartGroups(1)
    chartgroup.HasSeriesLines = true
    chartgroup.SeriesLines.Border.Color = RGB(0, 0, 255)
}
```

```javascript
/*本示例将图表工作表 Chart1 上第一个图表组的系列线边框的粗细设置为粗。*/
function test() {
    let chartgroup = Application.Charts.Item("Chart1").ChartObjects(1).Chart.ChartGroups(1)
    chartgroup.SeriesLines.Border.Weight = xlThick
}
```
