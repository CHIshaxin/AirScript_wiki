---
title: "Legend (对象)"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/Legend/Legend%20%E5%AF%B9%E8%B1%A1.html
updated: 1789700916000
hash: Bv6cuI-s
---

# Legend (对象)

代表图标中的图例。每个图表只能有一个图例。

## 说明

**[Legend](./../Legend/Legend 对象.html)** 对象包含一个或多个 **[LegendEntry](./../LegendEntry/LegendEntry 对象.html)** 对象；每个 **[LegendEntry](./../LegendEntry/LegendEntry 对象.html)** 对象都包含一个 **[LegendKey](./../LegendKey/LegendKey 对象.html)** 对象。

除非 **HasLegend** 属性是 **True**，否则无法看到图表的图例。如果该属性为 **False**， **[Legend](./../Legend/Legend 对象.html)** 对象的属性和方法将会失败。

## 示例

```javascript
/*下例将第一张工作表上嵌入图表一的图例字形设置为加粗。*/
function test() {
    Application.Worksheets.Item(1).ChartObjects(1).Chart.Legend.Font.Bold = true
}
```

```javascript
/*此示例删除图表工作表 Chart1 上的第一个图表中的图例。*/
function test() {
    Application.Charts.Item("Chart1").ChartObjects(1).Chart.Legend.Delete()
}
```
