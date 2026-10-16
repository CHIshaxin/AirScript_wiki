---
title: "HasRadarAxisLabels (属性)"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/ChartGroup/%E5%B1%9E%E6%80%A7/HasRadarAxisLabels%20%E5%B1%9E%E6%80%A7.html
updated: 1789700916000
hash: Bs0bmTkR
---

# HasRadarAxisLabels (属性)

如果雷达图有坐标轴标签，则该值为 **True**。仅适用于雷达图。**Boolean** 类型，可读写。

## 示例

```javascript
/*本示例显示图表工作表 Chart1 中第一个图表组的雷达图坐标轴标签，并对这些标签的颜色进行设置。本示例应在雷达图上运行。*/
function test() {
    let chartgroup = Application.Charts.Item("Chart1").ChartObjects(1).Chart.ChartGroups(1)
    chartgroup.HasRadarAxisLabels = true
    chartgroup.RadarAxisLabels.Font.ColorIndex = 3
}
```

```javascript
/*本示例判断工作表 Sheet1 第一个内嵌图表中第一个图表组是否含有雷达图坐标轴标签，并通知用户。本示例应在雷达图上运行。*/
function test() {
    let chartgroup = Application.Worksheets.Item("Sheet1").ChartObjects(1).Chart.ChartGroups(1)
    if (chartgroup.HasRadarAxisLabels) {
        console.log("图表组含有坐标轴标签")
    } else {
        console.log("图表组不含有坐标轴标签")
    }
}
```
