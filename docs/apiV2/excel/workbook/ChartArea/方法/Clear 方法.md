---
title: "Clear 方法"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/ChartArea/%E6%96%B9%E6%B3%95/Clear%20%E6%96%B9%E6%B3%95.html
---
# Clear (方法)

清除整个对象。

## 示例

```javascript
/*此示例清除图表工作表 Chart1 中的图表区域（图表数据和格式设置）。*/
function test() {
    let chartarea = Application.Charts.Item("Chart1").ChartArea
    chartarea.Clear()
}
```

```javascript
/*此示例清除工作表 Sheet1 第一个内嵌图表的图表区域（图表数据和格式设置）。*/
function test() {
    let chartarea = Application.Worksheets.Item("Sheet1").ChartObjects(1).Chart.ChartArea
    chartarea.Clear()
}
```
