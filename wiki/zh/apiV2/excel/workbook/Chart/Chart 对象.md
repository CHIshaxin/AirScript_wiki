---
title: "Chart (对象)"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/Chart/Chart%20%E5%AF%B9%E8%B1%A1.html
updated: 1789700916000
hash: D7E3Zssg
---

# Chart (对象)

代表工作簿中的图表。

## 说明

此图表既可以是嵌入的图表（包含在 **[ChartObject](./../ChartObject/ChartObject 对象.html)** 对象中），也可以是单独的图表工作表。

示例部分中描述了以下用于返回 **[Chart](./../Chart/Chart 对象.html)** 对象的属性和方法：

- **[Charts](./../Charts/Charts 对象.html)** 方法
- **ActiveChart** 属性
- **ActiveSheet** 属性

**[Charts](./../Charts/Charts 对象.html)** 集合包含工作簿中每个图表工作表的 **[Chart](./../Chart/Chart 对象.html)** 对象。使用 **[Charts](./../Charts/Charts 对象.html)** ( *index*) 可以返回单个 **[Chart](./../Chart/Chart 对象.html)** 对象，其中 index 为图表工作表的索引号或名称。图表索引号表示图表工作表在工作簿标签栏上的位置。*Charts(1)* 是工作簿中第一个（最左边的）图表；*Charts(Charts.Count)* 是最后一个（最右边的）图表。所有图表工作表均包括在索引计数中，即便是隐藏工作表也是如此。图表工作表名称显示在图表工作簿标签上。您可以使用 **Name** 属性设置或返回图表名称。

## 示例

```javascript
/*本示例将图表工作表 Chart1 第一个数据系列设置为红色。*/
function test() {
    let chart = Application.Charts.Item("Chart1").ChartObjects(1).Chart
    chart.SeriesCollection(1).Format.Fill.ForeColor.RGB = rgbRed
}
```

```javascript
/*本示例将图表工作表 Sales 移至活动工作簿的尾部。*/
function test() {
    let chart = Application.Charts.Item("Sales")
    chart.Move(null, Sheets.Item(Sheets.Count))
}
```

**[Chart](./../Chart/Chart 对象.html)** 对象也是 **[Sheets](./../Sheets/Sheets 对象.html)** 集合的成员，此集合包含工作簿中的所有工作表（图表工作表和工作表）。使用 **Sheets**(*index*) 可以返回单个工作表，其中 *index* 是工作表索引号或名称。

当图表是活动对象时，您可以使用 **ActiveChart** 属性引用它。如果用户选择了图表工作表，或者用 **[Chart](./../Chart/Chart 对象.html)** 对象的 **Activate** 方法或 **[ChartObject](./../ChartObject/ChartObject 对象.html)** 对象的 **Activate** 方法激活了它，则该图表工作表处于活动状态。

```javascript
/*本示例将图表工作表 Chart1 图表类型修改为折线图，并将图表标题修改为“January Sales”。*/
function test() {
    let chart = Application.Charts.Item("Chart1").ChartObjects(1).Chart
    chart.ChartType = xlLine
    chart.HasTitle = true
    chart.ChartTitle.Text = "January Sales"
}
```

如果用户选择了嵌入图表，或者用 **Activate** 方法激活了包含该嵌入图表的 **[ChartObject](./../ChartObject/ChartObject 对象.html)** 对象，则该嵌入图表处于活动状态。通过使用 **ActiveChart** 属性，您可以编写能够引用嵌入图表或图表工作表（视哪一个处于活动状态而定）的 JavaScript 代码。

```javascript
/*本示例将工作表 Sheet1 中第一个内嵌图表的图表类型修改为折线图，然后将图表标题修改为“January Sales”。*/
function test() {
    let chart = Application.Worksheets.Item("Sheet1").ChartObjects(1).Chart
    chart.ChartType = xlLine
    chart.HasTitle = true
    chart.ChartTitle.Text = "January Sales"
}
```

当图表工作表为活动工作表时，可以使用 **ActiveSheet** 属性来引用它。

```javascript
/*本示例将 Chart1 图表工作表中系列 1 的内部颜色设置为蓝色。*/
function test() {
    let chart = Application.Charts.Item("chart1").ChartObjects(1).Chart
    chart.SeriesCollection(1).Format.Fill.ForeColor.RGB = rgbBlue
}
```
