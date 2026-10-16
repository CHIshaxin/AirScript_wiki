---
title: "Left 属性"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/ChartObjects/%E5%B1%9E%E6%80%A7/Left%20%E5%B1%9E%E6%80%A7.html
---
# Left (属性)

返回或设置 **Double** 值，它代表从对象左边缘到工作表的 A 列左边缘或图表上的图表区左边缘的距离（以磅为单位）。

## 示例

```javascript
/*本示例将工作表 Sheet1 上所有图表到工作表A列的距离设置为 300 磅。*/
function test() {
    let chartobjects = Application.Sheets.Item("Sheet1").ChartObjects()
    chartobjects.Left = 300
}
```

```javascript
/*本示例显示图表工作表 Chart1 上图表到工作表A列的距离。*/
function test() {
    let chartobjects = Application.Charts.Item("Chart1").ChartObjects()
    console.log(`图表工作表上图表距A列：${chartobjects.Left} 磅。`)
}
```
