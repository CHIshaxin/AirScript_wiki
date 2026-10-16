---
title: "ChartStyle (属性)"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/Chart/%E5%B1%9E%E6%80%A7/ChartStyle%20%E5%B1%9E%E6%80%A7.html
updated: 1789700916000
hash: D-dwl-5S
---

# ChartStyle (属性)

返回或设置图表的图表样式。可读/写 **Variant** 类型。

## 说明

您可以使用介于 1 到 48 之间的数字设置图表样式。

## 示例

```javascript
/*本示例显示活动工作表上的第一个嵌入式图表的样式。*/
function test() {
    console.log(ActiveSheet.ChartObjects(1).Chart.ChartStyle)
}
```

```javascript
/*本示例修改图表工作表 Chart1 的图表样式。*/
function test() {
    let chart = Application.Charts.Item("Chart1")
    chart.ChartStyle = 4
}
```
