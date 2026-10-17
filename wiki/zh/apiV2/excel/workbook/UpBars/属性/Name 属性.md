---
title: "Name (属性)"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/UpBars/%E5%B1%9E%E6%80%A7/Name%20%E5%B1%9E%E6%80%A7.html
updated: 1789700916000
hash: BtLgNTXJ
---

# Name (属性)

返回一个 **String** 值，它代表对象的名称。

## 示例

```javascript
/*本示例显示图表工作表 Chart1 中图表的第一个图表组的涨柱线的名称。*/
function test() {
    let upbars = Application.Charts.Item("Chart1").ChartObjects(1).Chart.ChartGroups(1).UpBars
    console.log(upbars.Name)
}
```

```javascript
/*本示例显示工作表 Sheet1 中第一个图表的第一个图表组的涨柱线的名称。*/
function test() {
    let upbars = Application.Worksheets.Item("Sheet1").ChartObjects(1).Chart.ChartGroups(1).UpBars
    console.log(`涨柱线的名称：${upbars.Name}`)
}
```
