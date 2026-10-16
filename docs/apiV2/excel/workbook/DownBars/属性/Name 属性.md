---
title: "Name 属性"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/DownBars/%E5%B1%9E%E6%80%A7/Name%20%E5%B1%9E%E6%80%A7.html
---
# Name (属性)

返回一个 **String** 值，它代表对象的名称。

## 示例

```javascript
/*本示例显示图表工作表 Chart1 中图表的第一个图表组的跌柱线的名称。*/
function test() {
    let downbars = Application.Charts.Item("Chart1").ChartObjects(1).Chart.ChartGroups(1).DownBars
    console.log(downbars.Name)
}
```

```javascript
/*本示例显示工作表 Sheet1 中第一个图表的第一个图表组的跌柱线的名称。*/
function test() {
    let downbars = Application.Worksheets.Item("Sheet1").ChartObjects(1).Chart.ChartGroups(1).DownBars
    console.log(`跌柱线的名称：${downbars.Name}`)
}
```
