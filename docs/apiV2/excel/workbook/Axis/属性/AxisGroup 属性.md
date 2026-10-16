---
title: "AxisGroup 属性"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/Axis/%E5%B1%9E%E6%80%A7/AxisGroup%20%E5%B1%9E%E6%80%A7.html
---
# AxisGroup (属性)

返回指定轴的组。只读

## 返回值

XlAxisGroup

## 示例

```javascript
/*本示例显示图表工作表 Chart1 的类别轴是否为主轴组。*/
function test() {
    let axis = Application.Charts.Item("Chart1").ChartObjects(1).Chart.Axes(xlCategory)
    console.log(axis.AxisGroup == xlPrimary)
}
```

```javascript
/*本示例判断工作表 Sheet1 中第一张图表的类别轴是否为主轴组，并通知用户。*/
function test() {
    let axis = Application.Sheets.Item("Sheet1").ChartObjects(1).Chart.Axes(xlCategory)
    if (axis.AxisGroup == xlPrimary) {
        console.log("类别轴是主轴组")
    } else if (axis.AxisGroup == xlSecondary) {
        console.log("类别轴是辅助轴组")
    }
}
```
