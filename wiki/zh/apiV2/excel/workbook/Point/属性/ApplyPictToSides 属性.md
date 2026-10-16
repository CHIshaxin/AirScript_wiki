---
title: "ApplyPictToSides (属性)"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/Point/%E5%B1%9E%E6%80%A7/ApplyPictToSides%20%E5%B1%9E%E6%80%A7.html
updated: 1789700916000
hash: DQunywPL
---

# ApplyPictToSides (属性)

如果将图片应用于系列中某个数据点或所有数据点的旁边，则为 **True**。**Boolean** 类型，可读写。

## 示例

```javascript
/*此示例将图片应用于活动工作表第一个图表上第三个系列中第三个数据点旁边。图片必须已应用于该系列（设置此属性将更改图片的方向）。*/
function test() {
    Application.ActiveSheet.ChartObjects(1).Activate()
    ActiveChart.SeriesCollection(3).Points(3).ApplyPictToSides = true
}
```

```javascript
/*此示例显示图片是否应用于图表工作表 Chart1 上第二个系列的第一个数据点旁边。*/
function test() {
    let point = Application.Charts.Item("Chart1").ChartObjects(1).Chart.SeriesCollection(2).Points(1)
    console.log(point.ApplyPictToSides)
}
```
