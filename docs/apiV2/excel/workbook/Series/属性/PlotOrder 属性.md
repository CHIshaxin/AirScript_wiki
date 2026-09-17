---
title: "PlotOrder 属性"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/Series/%E5%B1%9E%E6%80%A7/PlotOrder%20%E5%B1%9E%E6%80%A7.html
---
# PlotOrder (属性)

返回或设置图表组中选定数据系列的绘制顺序。**Long** 类型，可读写。

## 说明

只能设置图表组内的绘制顺序（在图表中包含若干图表类型的情况下，不能设置整个图表的绘制顺序）。图表组指具有相同图表类型和子类型的数据系列的集合。

更改某个系列的绘制顺序将导致同一图表组内其他数据系列的绘制顺序的必要调整。

## 示例

```javascript
/*本示例使 Chart1 中的数据系列二在绘制时第三个出现。本示例应在至少包含三个数据系列的二维柱形图上运行。*/
function test() {
    Application.Charts.Item("Chart1").ChartObjects(1).Chart.ChartGroups(1).SeriesCollection(2).PlotOrder = 3
}
```

```javascript
/*此示例显示活动工作表中第一个图表的第二个数据系列在绘制时第几个出现。*/
function test() {
    console.log(ActiveSheet.ChartObjects(1).Chart.SeriesCollection(2).PlotOrder)
}
```
