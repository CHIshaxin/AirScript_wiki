---
title: "ApplyDataLabels 方法"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/Chart/%E6%96%B9%E6%B3%95/ApplyDataLabels%20%E6%96%B9%E6%B3%95.html
---
# ApplyDataLabels (方法)

将数据标签应用到图表中的所有系列。

## 说明

使用 **Separator** 参数时。如果使用字符串，您将得到一个字符串作为分隔符。如果您使用 **xlDataLabelSeparatorDefault** (= 1)，您将获得默认的数据标签分隔符，它是逗号或换行符，具体取决于数据标签。

| **XlDataLabelsType** 可为以下 **XlDataLabelsType** 常量之一。 |
| --- |
| **xlDataLabelsShowBubbleSizes**。数据标签的气泡尺寸。 |
| **xlDataLabelsShowLabelAndPercent**。占总数的百分比及数据点所属的分类。仅用于饼图和圆环图。 |
| **xlDataLabelsShowPercent**。占总数的百分比。仅用于饼图和圆环图。 |
| **xlDataLabelsShowLabel**。数据点所属的分类。 |
| **xlDataLabelsShowNone**。无数据标签。 |
| **xlDataLabelsShowValue**。默认。数据点的值（若未指定此参数，则假定为此值）。 |

## 参数

| 属性 | 数据类型 | 必填 | 说明 |
| --- | --- | --- | --- |
| Type | Default<Kso.XlDataLabelsType, 2> | 可选 | 要应用的数据标签的类型。 |
| LegendKey | any | 可选 | 如果为 True，则在数据点旁边显示图例项标示。默认值为 False。 |
| AutoText | any | 可选 | 如果对象根据内容自动生成相应的文字，则为 True。 |
| HasLeaderLines | any | 可选 | 对于 Chart 和 Series 对象，如果数据系列有引导线，则为 True。 |
| ShowSeriesName | any | 可选 | 传递布尔值以启用或禁用数据标签的系列名称。 |
| ShowCategoryName | any | 可选 | 传递布尔值以启用或禁用数据标签的分类名称。 |
| ShowValue | any | 可选 | 传递布尔值以启用或禁用数据标签的值。 |
| ShowPercentage | any | 可选 | 传递布尔值以启用或禁用数据标签的百分比。 |
| ShowBubbleSize | any | 可选 | 传递布尔值以启用或禁用数据标签的气泡大小。 |
| Separator | any | 可选 | 数据标签的分隔符。 |

## 示例

```javascript
/*本示例先清除活动工作表中第一个图表的数据标签，再对该图表的所有系列应用百分比及类别标签，并显示图例项标示以及使用分号作为分隔符。*/
function test() {
    ActiveSheet.ChartObjects(1).Chart.ApplyDataLabels(xlDataLabelsShowNone)
    ActiveSheet.ChartObjects(1).Chart.ApplyDataLabels(xlDataLabelsShowLabelAndPercent, true, null, null, true, false, null, null, null, ";")
}
```

```javascript
/*本示例先清除第一张工作表中第一个图表的数据标签，再对该图表的所有系列应用数据标签，显示数据标签的系列名与类别名。*/
function test() {
    let chart = Worksheets.Item(1).ChartObjects(1).Chart
    chart.ApplyDataLabels(xlDataLabelsShowNone)
    chart.ApplyDataLabels(null, false, null, null, true, true)
}
```
