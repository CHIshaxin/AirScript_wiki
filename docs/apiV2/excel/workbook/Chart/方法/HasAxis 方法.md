---
title: "HasAxis 方法"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/Chart/%E6%96%B9%E6%B3%95/HasAxis%20%E6%96%B9%E6%B3%95.html
---
# HasAxis (方法)

返回或设置图表上显示的坐标轴。**Variant** 类型，可读写。

## 说明

设置该属性时，必须至少输入一个参数的值。

如果更改了图表类型或 **Axis.AxisGroup**、**Chart.AxisGroup** 或 **Series.AxisGroup** 属性，ET 可能要创建或删除坐标轴。

## 参数

| 属性 | 数据类型 | 必填 | 说明 |
| --- | --- | --- | --- |
| Index1 | any | 可选 | 坐标轴类型。系列坐标轴仅适用于三维图表。可以是 **XlAxisType** 常量之一。 |
| Index2 | any | 可选 | 坐标轴组。三维图表只有一组坐标轴。可以是 **XlAxisGroup** 常量之一。 |

## 示例

```javascript
/*本示例显示图表工作表 Chart1 中的主数值轴。*/
function test() {
    let chart = Application.Charts.Item("Chart1").ChartObjects(1).Chart
    chart.HasAxis(xlValue, xlPrimary, true)
}
```

```javascript
/*本示例设置 Sheet1 工作表第一个内嵌图表的主数值轴为不可见。*/
function test() {
    let chart = Application.Worksheets.Item("Sheet1").ChartObjects(1).Chart
    chart.HasAxis(xlValue, xlPrimary, false)
}
```
