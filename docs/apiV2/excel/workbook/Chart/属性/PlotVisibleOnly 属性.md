---
title: "PlotVisibleOnly 属性"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/Chart/%E5%B1%9E%E6%80%A7/PlotVisibleOnly%20%E5%B1%9E%E6%80%A7.html
---
# PlotVisibleOnly (属性)

如果仅绘制可见单元格，则该值为 **True**。如果可见单元格和隐藏单元格都绘制，则该值为 **False**。**Boolean** 类型，可读写。

## 示例

```javascript
/*本示例设置图表工作表 Chart1 为仅绘制可见单元格。*/
function test() {
    Application.Charts.Item("Chart1").PlotVisibleOnly = true
}
```

```javascript
/*本示例判断工作表Sheet1第一个内嵌图表绘制单元格的方法，并通知用户。*/
function test() {
    let chart = Application.Worksheets.Item("Sheet1").ChartObjects(1).Chart
    if (chart.PlotVisibleOnly) {
        console.log("只绘制可见单元格")
    } else {
        console.log("绘制可见和隐藏单元格")
    }
}
```
