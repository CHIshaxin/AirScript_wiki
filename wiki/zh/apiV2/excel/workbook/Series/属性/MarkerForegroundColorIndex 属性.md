---
title: "MarkerForegroundColorIndex (属性)"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/Series/%E5%B1%9E%E6%80%A7/MarkerForegroundColorIndex%20%E5%B1%9E%E6%80%A7.html
updated: 1789700916000
hash: 2tr-70Yy
---

# MarkerForegroundColorIndex (属性)

返回或设置数据标志的前景色，表示为当前调色板中的索引或下列 **XlColorIndex** 常量之一：**xlColorIndexAutomatic** 或 **xlColorIndexNone**。仅适用于折线图、散点图和雷达图。**Long** 类型，可读写。

## 示例

```javascript
/*此示例通过前景色的索引值将图表工作表 Chart1 上的第二个数据系列数据标志的前景色设置为蓝色。*/
function test() {
    let series2 = Application.Charts.Item("Chart1").ChartObjects(1).Chart.SeriesCollection(2)
    series2.MarkerForegroundColorIndex = 5
}
```

```javascript
/*此示例显示活动工作表上的第一个图表的第一个系列数据标志的前景色是否为无颜色。*/
function test() {
    console.log(ActiveSheet.ChartObjects(1).Chart.SeriesCollection(1).MarkerForegroundColorIndex == xlColorIndexNone)
}
```
