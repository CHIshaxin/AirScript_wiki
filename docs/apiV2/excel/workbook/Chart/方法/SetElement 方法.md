---
title: "SetElement 方法"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/Chart/%E6%96%B9%E6%B3%95/SetElement%20%E6%96%B9%E6%B3%95.html
---
# SetElement (方法)

设置图表上的图表元素。可读/写 **MsoChartElementType** 类型。

## 说明

对于图表，\*\*“布局”\*\*选项卡中的以下命令对应于 **SetElement** 方法：

- \*\*“标签”\*\*组中的全部内容。
- \*\*“坐标轴”\*\*组中的全部内容。
- \*\*“分析”\*\*组中的全部内容。
- **“绘图区”**、\*\*“图表背景墙”**和**“图表基底”\*\*按钮。

**MsoChartElementType** 是常量枚举，引用上述所有命令。

## 参数

| 属性 | 数据类型 | 必填 | 说明 |
| --- | --- | --- | --- |
| Element | Kso.MsoChartElementType | 必填 | 指定图表元素类型。 |

## 示例

```javascript
/*此示例使用不同常量值向活动图表设置图表元素。*/
function test() {
    let chart = Application.Charts.Item("Chart1").ChartObjects(1).Chart
    chart.Axes(xlValue).MajorGridlines.Select()
    chart.SetElement(msoElementChartTitleCenteredOverlay)
    chart.SetElement(msoElementPrimaryCategoryGridLinesMinor)
    chart.SetElement(msoElementChartTitleNone)
}
```

```javascript
/*此示例在图表工作表 Chart1 上居中显示数据标签。*/
function test() {
    let chart = Application.Charts.Item("Chart1").ChartObjects(1).Chart
    chart.SetElement(msoElementDataLabelCenter)
}
```
