---
title: "ErrorBar 方法"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/Series/%E6%96%B9%E6%B3%95/ErrorBar%20%E6%96%B9%E6%B3%95.html
---
# ErrorBar (方法)

将误差线应用于系列。**Variant** 类型。

## 参数

| 属性 | 数据类型 | 必填 | 说明 |
| --- | --- | --- | --- |
| Direction | Kso.XlErrorBarDirection | 必填 | 误差线方向。 |
| Include | Kso.XlErrorBarInclude | 必填 | 要包含的误差线部分。 |
| Type | Kso.XlErrorBarType | 必填 | 误差线类型。 |
| Amount | any | 可选 | 误差量。当 Type 为 xlErrorBarTypeCustom 时只用于正误差量。 |
| MinusValues | any | 可选 | 当 Type 为 xlErrorBarTypeCustom 时，为负误差量。 |

## 返回值

Variant

## 示例

```javascript
/*本示例显示 Chart1 中的第一个数据系列在 Y 方向上的标准误差线。误差线是正负两个方向的。本示例应在二维折线图上运行。*/
function test() {
    Application.Charts.Item("Chart1").ChartObjects(1).Chart.SeriesCollection().Item(1).ErrorBar(xlY, xlErrorBarIncludeBoth, xlErrorBarTypeStError)
}
```

```javascript
/*此示例显示活动工作表中的第一个数据系列在 Y 方向上的百分比类型误差线。误差线是正方向的,误差量为10。本示例应在二维折线图上运行。*/
function test() {
    ActiveSheet.ChartObjects(1).Chart.SeriesCollection(1).ErrorBar(xlY, xlErrorBarIncludePlusValues, xlErrorBarTypePercent, 10)
}
```
