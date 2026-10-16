---
title: "ErrorBars 属性"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/Series/%E5%B1%9E%E6%80%A7/ErrorBars%20%E5%B1%9E%E6%80%A7.html
---
# ErrorBars (属性)

返回一个 **[ErrorBars](./../../ErrorBars/ErrorBars 对象.html)** 对象，该对象表示数据系列中的误差线。只读。

## 示例

```javascript
/*本示例对 Chart1 中第一个系列的误差线颜色进行设置。本示例应在其第一个系列有误差线的二维折线图上运行。*/
function test() {
    let Ser = Application.Charts.Item("Chart1").ChartObjects(1).Chart.SeriesCollection(1)
    Ser.ErrorBars.Border.ColorIndex = 8
}
```

```javascript
/*此示例显示活动工作表中的第一个数据系列的误差线名称。*/
function test() {
    console.log(ActiveSheet.ChartObjects(1).Chart.SeriesCollection(1).ErrorBars.Name)
}
```
