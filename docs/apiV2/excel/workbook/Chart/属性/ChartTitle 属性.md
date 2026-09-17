---
title: "ChartTitle 属性"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/Chart/%E5%B1%9E%E6%80%A7/ChartTitle%20%E5%B1%9E%E6%80%A7.html
---
# ChartTitle (属性)

返回一个 **[ChartTitle](./../../ChartTitle/ChartTitle 对象.html)** 对象，该对象表示指定图表的标题。只读。

## 示例

```javascript
/*本示例设置图表工作表 Chart1 的标题文本为“First Quarter Sales”。*/
function test() {
    let charts = Application.Charts.Item("Chart1")
    charts.HasTitle = true
    charts.ChartTitle.Text = "First Quarter Sales"
}
```

```javascript
/*本示例删除活动工作表第一个嵌入式图表的图表标题。*/
function test() {
    let charts = ActiveSheet.ChartObjects(1).Chart
    charts.ChartTitle.Delete()
}
```
