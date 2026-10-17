---
title: "TickLabels (对象)"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/TickLabels/TickLabels%20%E5%AF%B9%E8%B1%A1.html
updated: 1789700916000
hash: Dw0N_B1x
---

# TickLabels (对象)

代表图表坐标轴上刻度线的刻度线标志。

## 说明

此对象不是集合。没有代表单个刻度线标志的对象；您必须将所有刻度线标志作为一个单位返回。

分类轴的刻度线标签文字来自图表中相应分类的名称。分类轴的默认刻度线标签文字表示该分类相对于该坐标轴最左端偏移量的数字。要更改分类轴刻度线标签间不带标签的刻度线数量，您必须更改分类轴的 **TickLabelSpacing** 属性。

数值轴的刻度线标志文字的计算，是基于数值轴的 **MajorUnit**、**MinimumScale** 和 **MaximumScale** 属性。要更改数值的轴刻度线标签文本，您必须更改这些属性的值。

使用 **TickLabels** 属性可返回 **[TickLabels](./../TickLabels/TickLabels 对象.html)** 对象。

## 示例

```javascript
/*下例设置 Sheet1 上嵌入式图表一中数值轴上刻度线标志的数字格式。*/
function test() {
    Worksheets.Item("Sheet1").ChartObjects(1).Chart.Axes(xlValue).TickLabels.NumberFormat = "0.00"
}
```

```javascript
/*以下示例删除 Chart1 上分类轴的刻度线标志。*/
function test() {
    let tickLabels = Application.Charts.Item("Chart1").ChartObjects(1).Chart.Axes(xlCategory).TickLabels
    tickLabels.Delete()
}
```
