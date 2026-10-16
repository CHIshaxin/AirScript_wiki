---
title: "MarkerBackgroundColor 属性"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/Series/%E5%B1%9E%E6%80%A7/MarkerBackgroundColor%20%E5%B1%9E%E6%80%A7.html
---
# MarkerBackgroundColor (属性)

将数据标志的背景色设置为 RGB 值或返回对应的颜色索引值。仅适用于折线图、散点图和雷达图。**Long**型，可读写。

## 示例

```javascript
/*此示例将图表工作表 Chart1 上的第二个数据系列数据标志的背景色设置为红色。*/
function test() {
    let series2 = Application.Charts.Item("Chart1").ChartObjects(1).Chart.SeriesCollection(2)
    series2.MarkerBackgroundColor = RGB(255, 0, 0)
}
```

```javascript
/*此示例显示活动工作表上的第一个图表的第一个系列数据标志的背景色是否为绿色。*/
function test() {
    console.log(ActiveSheet.ChartObjects(1).Chart.SeriesCollection(1).MarkerBackgroundColor == RGB(0, 255, 0))
}
```
