---
title: "RoundedCorners (属性)"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/ChartArea/%E5%B1%9E%E6%80%A7/RoundedCorners%20%E5%B1%9E%E6%80%A7.html
updated: 1789700916000
hash: D8_Y4ETb
---

# RoundedCorners (属性)

如果图表的图表区域使用圆角，则为 **True**。可读写 **Boolean** 类型。

## 示例

```javascript
/*此示例为工作表 Sheet1 上的第一个图表区添加圆角。*/
function test() {
    let chartarea = Application.Worksheets.Item("Sheet1").ChartObjects(1).Chart.ChartArea
    chartarea.RoundedCorners = true
}
```

```javascript
/*此示例显示图表工作表 Chart1 图表区是否为圆角。*/
function test() {
    let chartarea = Application.Charts.Item("Chart1").ChartObjects(1).Chart.ChartArea
    console.log(chartarea.RoundedCorners)
}
```
