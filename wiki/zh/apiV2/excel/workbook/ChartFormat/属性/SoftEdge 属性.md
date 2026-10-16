---
title: "SoftEdge (属性)"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/ChartFormat/%E5%B1%9E%E6%80%A7/SoftEdge%20%E5%B1%9E%E6%80%A7.html
updated: 1789700916000
hash: 11rtEacA
---

# SoftEdge (属性)

返回指定图表的 **SoftEdgeFormat** 对象，该对象包含图表的柔化边缘格式设置属性。只读。

## 示例

```javascript
/*此示例显示图表工作表 Chart1 图表区柔化边缘效果的大小（以磅为单位）。*/
function test() {
    let chartformat = Application.Charts.Item("Chart1").ChartObjects(1).Chart.ChartArea.Format
    console.log(chartformat.SoftEdge.Radius)
}
```

```javascript
/*此示例设置工作表 Sheet1 第一个内嵌图表区柔化边缘效果的大小为 0.5 磅。*/
function test() {
    let chartformat = Application.Worksheets.Item("Sheet1").ChartObjects(1).Chart.ChartArea.Format
    chartformat.SoftEdge.Radius = 0.5
}
```
