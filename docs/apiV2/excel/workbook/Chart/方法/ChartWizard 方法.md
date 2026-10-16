---
title: "ChartWizard 方法"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/Chart/%E6%96%B9%E6%B3%95/ChartWizard%20%E6%96%B9%E6%B3%95.html
---
# ChartWizard (方法)

修改给定图表的属性。可使用本方法快速设置图表的格式，而不必逐个设置所有属性。本方法是非交互式的，并且仅更改指定的属性。

## 说明

如果省略 *Source*，并且选定内容不是活动工作表中的嵌入图表或者活动工作表不是当前存在的图表，则该方法失效并产生错误。

## 参数

| 属性 | 数据类型 | 必填 | 说明 |
| --- | --- | --- | --- |
| Source | any | 可选 | 包含新图表源数据的区域。如果省略本参数，ET 将编辑活动图表工作表或活动工作表上处于选定状态的图表。 |
| Gallery | any | 可选 | 用于指定图表类型的 XlChartType 的常量之一。 |
| Format | any | 可选 | 内置自动套用格式的选项编号。可为从 1 到 10 的数字，其取值取决于库的类型。如果省略此参数，ET 将根据库的类型和数据源选择默认值。 |
| PlotBy | any | 可选 | 指定每个系列的数据是来自行还是来自列。可以是以下 XlRowCol 常量之一：xlRows 或 xlColumns。 |
| CategoryLabels | any | 可选 | 指定包含分类标签的源范围内的行数或列数的整数。合法值为从 0（零）至小于相应分类或系列的最大个数间的某一数字。 |
| SeriesLabels | any | 可选 | 指定包含系列标志的源范围内的行数或列数的整数。合法值为从 0（零）至小于相应分类或系列的最大个数间的某一数字。 |
| HasLegend | any | 可选 | 若要包括图例，则为 True。 |
| Title | any | 可选 | 图表标题文字。 |
| CategoryTitle | any | 可选 | 分类轴标题文字。 |
| ValueTitle | any | 可选 | 数值轴标题文字。 |
| ExtraTitle | any | 可选 | 三维图表的系列轴标题，或二维图表的次数值轴标题。 |

## 示例

```javascript
/*本示例重新设置图表工作表 Chart1 的格式，将其改为折线图，添加图例，并添加分类轴标题和数值轴标题。*/
function test() {
    Application.Charts.Item("Chart1").ChartWizard(null, xlLine, null, null, null, null, true, null, "Year", "Sales")
}
```

```javascript
/*本示例设置图表工作表 Chart1 的源数据区域为第一张工作表区域A1:C5，并将图表工作表设置为百分比堆积柱形图第三个版式，添加图例、分类轴标题和数值轴标题。*/
function test() {
    let chart = Application.Charts.Item("Chart1")
    chart.ChartWizard(Worksheets.Item(1).Range("A1:C5"), xlColumnStacked100, 3, xlColumns, null, null, true, null, "Age", "Year")
}
```
