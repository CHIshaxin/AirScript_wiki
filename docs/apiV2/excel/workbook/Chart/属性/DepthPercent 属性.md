---
title: "DepthPercent 属性"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/Chart/%E5%B1%9E%E6%80%A7/DepthPercent%20%E5%B1%9E%E6%80%A7.html
---
# DepthPercent (属性)

返回或设置三维图表的深度，以图表宽度的百分比表示（有效范围从 20% 到 2000%）。**Long** 类型，可读写。

## 示例

```javascript
/*本示例将图表工作表 Chart1 的深度设置为其宽度的 50%。*/
function test() {
    //本示例应在三维图表上运行 （DepthPercent 属性在二维图表中无效）。
    Application.Charts.Item("Chart1").DepthPercent = 50
}
```

```javascript
/*本示例显示第一张工作表第一张内嵌图表的深度。*/
function test() {
    //本示例应在三维图表上运行（DepthPercent 属性在二维图表中无效）。
    let chart = Application.Worksheets.Item(1).ChartObjects(1).Chart
    console.log(chart.DepthPercent)
}
```
