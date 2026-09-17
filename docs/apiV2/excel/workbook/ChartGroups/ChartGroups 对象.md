---
title: "ChartGroups 对象"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/ChartGroups/ChartGroups%20%E5%AF%B9%E8%B1%A1.html
---
# ChartGroups (对象)

代表图表中用同一格式绘制的一个或多个数据系列。

## 说明

**[ChartGroups](./../ChartGroups/ChartGroups 对象.html)** 集合是指定图表中的所有 **[ChartGroup](./../ChartGroup/ChartGroup 对象.html)** 对象的集合。每张图表包含一个或多个图表组，每个图表组包含一个或多个数据系列，每个数据系列包含一个或多个数据点。例如，单张图表可能既包含折线图图表组（其中包含所有用折线图格式绘制的数据系列），也包含条形图图表组（其中包含所有用条形图格式绘制的数据系列）。

使用 **ChartGroups** 方法可返回 **[ChartGroups](./../ChartGroups/ChartGroups 对象.html)** 集合。

## 示例

```javascript
/*本示例显示工作表 Sheet1 中第一个内嵌图表图表组的数量。*/
function test() {
    let  chartgroups = Application.Worksheets.Item("Sheet1").ChartObjects(1).Chart.ChartGroups()
    console.log(chartgroups.Count)
}
```

使用 **ChartGroups**(*index*)（其中 *index* 是图表组的索引号）可以返回单个 **[ChartGroup](./../ChartGroup/ChartGroup 对象.html)** 对象。

```javascript
/*本示例为图表工作表 Chart1 中第一个图表组添加垂直线。*/
function test() {
    let chartgroup = Application.Charts.Item("Chart1").ChartObjects(1).Chart.ChartGroups(1)
    chartgroup.HasDropLines = true
}
```

如果图表已被激活，可使用 **ActiveChart**：

```javascript
/*本示例激活图表工作表 Chart1 ，并为该表的第一个图表组添加垂直线。*/
function test() {
    Application.Charts.Item("Chart1").Activate()
    Application.ActiveSheet.ChartGroups(1).HasDropLines = true
}
```

因为当特定图表组所用的图表格式更改时，该图表组的索引号可能会更改，所以使用命名图表组快捷方法之一来返回特定的图表组会更加容易。**PieGroups** 方法返回图表中饼图图表组的集合，**LineGroups** 方法返回图表中折线图图表组的集合，依此类推。这些方法中的每一个都可以与索引号配合使用以返回单个 **[ChartGroup](./../ChartGroup/ChartGroup 对象.html)** 对象，或不指定索引号而返回 **[ChartGroups](./../ChartGroups/ChartGroups 对象.html)** 集合。
