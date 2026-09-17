---
title: "Charts 对象"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/Charts/Charts%20%E5%AF%B9%E8%B1%A1.html
---
# Charts (对象)

指定的或活动工作簿中所有图表工作表的集合。

## 示例

```javascript
/*本示例设置图表工作表 Chart1 中图表的配色方案。*/
function test() {
    let chart = Application.Charts.Item("Chart1")
    chart.ChartColor = 7
}
```

```javascript
/*本示例在活动工作簿所有图表工作表上添加水平分页符。*/
function test() {
    Application.Charts.HPageBreaks.Add(Range("C3"))
}
```
