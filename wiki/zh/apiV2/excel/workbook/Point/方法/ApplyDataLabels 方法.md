---
title: "ApplyDataLabels (方法)"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/Point/%E6%96%B9%E6%B3%95/ApplyDataLabels%20%E6%96%B9%E6%B3%95.html
updated: 1789700916000
hash: DW5TUWKj
---

# ApplyDataLabels (方法)

对数据点应用数据标签。

## 说明

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
| Type | Default | 可选 | 要应用的数据标签的类型。 |
| LegendKey | any | 可选 | 如果为 **True**，则在数据点旁边显示图例项标示。默认值为 **False**。 |
| AutoText | any | 可选 | 如果对象根据内容自动生成相应的文字，则为 **True**。 |
| HasLeaderLines | any | 可选 | 对于 **Chart** 和 **[Series](./../../Series/Series 对象.html)** 对象，如果数据系列有引导线，则为 **True**。 |
| ShowSeriesName | any | 可选 | 传递布尔值以启用或禁用数据标签的系列名称。 |
| ShowCategoryName | any | 可选 | 传递布尔值以启用或禁用数据标签的分类名称。 |
| ShowValue | any | 可选 | 传递布尔值以启用或禁用数据标签的值。 |
| ShowPercentage | any | 可选 | 传递布尔值以启用或禁用数据标签的百分比。 |
| ShowBubbleSize | any | 可选 | 传递布尔值以启用或禁用数据标签的气泡大小。 |
| Separator | any | 可选 | 数据标签的分隔符。 |

## 示例

```javascript
/*此示例对图表工作表 Chart1 上第一个系列的第一个数据点应用分类标签*/
function test() {
    let point = Application.Charts.Item("Chart1").ChartObjects(1).Chart.SeriesCollection(1).Points(1)
    point.ApplyDataLabels(xlDataLabelsShowLabel)
}
```

```javascript
/*此示例对活动工作表第一个图表上第二个系列的第二个数据点应用绝对值标签，显示图例项标示并启用数据点的值。*/
function test() {
    Application.ActiveSheet.ChartObjects(1).Activate()
    ActiveChart.SeriesCollection(2).Points(2).ApplyDataLabels(xlDataLabelsShowBubbleSizes, true, null, null, null, null, true)
}
```
