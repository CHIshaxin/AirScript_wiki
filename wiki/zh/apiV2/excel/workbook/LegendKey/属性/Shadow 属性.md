---
title: "Shadow (属性)"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/LegendKey/%E5%B1%9E%E6%80%A7/Shadow%20%E5%B1%9E%E6%80%A7.html
updated: 1789700916000
hash: oXdp3698
---

# Shadow (属性)

返回或设置一个 **Boolean** 值，它确定对象是否有阴影。

## 示例

```javascript
/*下例显示 Sheet1 中的第一个图表的第二个图例项的图例标示是否有阴影。*/
function test() {
    let legend = Application.Worksheets.Item("Sheet1").ChartObjects(1).Chart.Legend
    console.log(legend.LegendEntries(2).LegendKey.Shadow)
}
```

```javascript
/*本示例将 Chart1 上第一个图例项的图例标示设置成有阴影。*/
function test() {
    Application.Charts.Item("Chart1").ChartObjects(1).Chart.Legend.LegendEntries(1).LegendKey.Shadow = true
}
```
