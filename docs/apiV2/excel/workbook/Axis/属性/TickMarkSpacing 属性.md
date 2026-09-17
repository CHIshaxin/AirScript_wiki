---
title: "TickMarkSpacing 属性"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/Axis/%E5%B1%9E%E6%80%A7/TickMarkSpacing%20%E5%B1%9E%E6%80%A7.html
---
# TickMarkSpacing (属性)

返回或设置刻度线之间的分类数或数据系列数。仅应用于分类轴和系列轴。可以是 1 到 31999 之间的一个数值。**Long** 类型，可读写。

## 说明

可用 **MajorUnit** 属性和 **MinorUnit** 属性设置数值轴上的刻度线间距。

## 示例

```javascript
/*本示例设置图表工作表 Chart1 中图表的分类轴上刻度线之间的分类数。*/
function test() {
    let axis = Application.Charts.Item("Chart1").ChartObjects(1).Chart.Axes(xlCategory)
    axis.TickMarkSpacing = 10
}
```

```javascript
/*本示例显示工作表 Sheet1 中第一张图表的分类轴上刻度线之间的分类数是否为 5。*/
function test() {
    let axis = Application.Sheets.Item("Sheet1").ChartObjects(1).Chart.Axes(xlCategory)
    console.log(axis.TickMarkSpacing == 5)
}
```
