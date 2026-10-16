---
title: "AxisGroup (属性)"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/Series/%E5%B1%9E%E6%80%A7/AxisGroup%20%E5%B1%9E%E6%80%A7.html
updated: 1789700916000
hash: DUABjQVB
---

# AxisGroup (属性)

返回或设置指定系列的组。可读/写。

## 返回值

XlAxisGroup

## 示例

```javascript
/*此示例将图表工作表 Chart1 上的第一个数据系列坐标轴组的类型设置为辅助轴组。*/
function test() {
    let series1 = Application.Charts.Item("Chart1").ChartObjects(1).Chart.SeriesCollection(1)
    series1.AxisGroup = xlSecondary
}
```

```javascript
/*此示例显示活动工作表上的第一个图表的第一个系列坐标轴组的类型是否为主轴组。*/
function test() {
    console.log(ActiveSheet.ChartObjects(1).Chart.SeriesCollection(1).AxisGroup == xlPrimary)
}
```
