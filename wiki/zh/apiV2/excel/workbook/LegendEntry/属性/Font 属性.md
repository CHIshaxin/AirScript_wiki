---
title: "Font (属性)"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/LegendEntry/%E5%B1%9E%E6%80%A7/Font%20%E5%B1%9E%E6%80%A7.html
updated: 1789700916000
hash: Da5g-kaE
---

# Font (属性)

返回一个 **[Font](./../../Font/Font 对象.html)** 对象，它代表指定对象的字体。

## 示例

```javascript
/*下例将 Sheet1 中的第一个图表的第二个图例项字体设置加粗。*/
function test() {
    let legend = Application.Worksheets.Item("Sheet1").ChartObjects(1).Chart.Legend
    legend.LegendEntries(2).Font.Bold = true
}
```

```javascript
/*下例将 Chart1 上第一个图例项字体颜色设置成绿色。*/
function test() {
    let legend = Application.Charts.Item("Chart1").ChartObjects(1).Chart.Legend
    legend.LegendEntries(1).Font.Color = RGB(0, 255, 0)
}
```
