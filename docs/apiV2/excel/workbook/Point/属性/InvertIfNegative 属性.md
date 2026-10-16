---
title: "InvertIfNegative 属性"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/Point/%E5%B1%9E%E6%80%A7/InvertIfNegative%20%E5%B1%9E%E6%80%A7.html
---
# InvertIfNegative (属性)

如果指定项与一个负数相对应时 Kingsoft WPS表格 就将其反色，则为 **True**。**Boolean** 类型，可读写。

## 示例

```javascript
/*此示例设置活动工作表第一个图表上第二个系列的第二个数据点的负值将以相反的颜色进行显示*/
function test() {
    Application.ActiveSheet.ChartObjects(1).Activate()
    ActiveChart.SeriesCollection(2).Points(2).InvertIfNegative = true
}
```

```javascript
/*此示例设置Chart1中第三个系列所有偶数位的数据点的负值将以相反的颜色进行显示。*/
function test() {
    let series = Application.Charts.Item("Chart1").ChartObjects(1).Chart.SeriesCollection(3)
    for (let i = 2; i <= series.Points().Count; i += 2) {
        series.Points(i).InvertIfNegative = true
    }
}
```
