---
title: "ChartGroup (对象)"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/ChartGroup/ChartGroup%20%E5%AF%B9%E8%B1%A1.html
updated: 1789700916000
hash: Cfw1OThX
---

# ChartGroup (对象)

代表图表中用同一格式绘制的一个或多个数据系列。

## 说明

一张图表包含一个或多个图表组，每个图表组包含一个或多个 **[Series](./../Series/Series 对象.html)** 对象，每个数据系列包含一个或多个 **[Points](./../Points/Points 对象.html)** 对象。例如，单张图表可能既包含折线图图表组（其中包含所有用折线图格式绘制的数据系列），也包含条形图图表组（其中包含所有用条形图格式绘制的数据系列）。 **[ChartGroup](./../ChartGroup/ChartGroup 对象.html)** 对象是 **[ChartGroups](./../ChartGroups/ChartGroups 对象.html)** 集合的成员。

使用 **ChartGroups**(*index*)（其中 *index* 是图表组的索引号）可以返回单个 **[ChartGroup](./../ChartGroup/ChartGroup 对象.html)** 对象。

因为当特定图表组所用的图表格式更改时，该图表组的索引号可能会更改，所以使用命名图表组快捷方法之一来返回特定的图表组会更加容易。**PieGroups** 方法返回图表中饼图图表组的集合，**LineGroups** 方法返回图表中折线图图表组的集合，依此类推。这些方法中的每一个都可以与索引号配合使用以返回单个 **[ChartGroup](./../ChartGroup/ChartGroup 对象.html)** 对象，或不指定索引号而返回 **[ChartGroups](./../ChartGroups/ChartGroups 对象.html)** 集合。

## 示例

```javascript
/*本示例开启图表工作表 Chart1 中第一个图表组的垂直线。*/
function test() {
    let chartgroup = Application.Charts.Item("Chart1").ChartObjects(1).Chart.ChartGroups(1)
    chartgroup.HasDropLines = true
}
```

如果图表已被激活，就可使用 **ActiveChart** 属性。

```javascript
/*本示例激活图表工作表 Chart1 ，并开启该图表中第一个图表组的垂直线。*/
function test() {
    Application.Charts.Item("Chart1").Activate()
    Application.ActiveSheet.ChartObjects(1).Chart.ChartGroups(1).HasDropLines = true
}
```
