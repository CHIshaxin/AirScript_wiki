---
title: "HiLoLines 对象"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/HiLoLines/HiLoLines%20%E5%AF%B9%E8%B1%A1.html
---
# HiLoLines (对象)

代表图表组中的高低点连线。

## 说明

高低点连线连接图表组内每一分类中的最高数据点和最底数据点。只有二维折线图组可以有高低点连线。此对象不是集合。没有代表单个高低点连线的对象；或者打开图表组中所有数据点的高低点连线，或者将其全部关闭。

如果 **HasHiLoLines** 属性是 **False**， **[HiLoLines](./../HiLoLines/HiLoLines 对象.html)** 对象的大多数属性都会被禁用。

使用 **HiLoLines** 属性可返回 **[HiLoLines](./../HiLoLines/HiLoLines 对象.html)** 对象。

## 示例

```javascript
/*下例使用 HasHiLowLines 属性将 HiLowLines 添加到工作表一上的嵌入图表一（该图表必须是折线图），并将高低点连线设置为蓝色。*/
function test() {
    Application.Worksheets.Item(1).ChartObjects(1).Activate()
    ActiveChart.ChartGroups(1).HasHiLoLines = true
    ActiveChart.ChartGroups(1).HiLoLines.Border.Color = RGB(0, 0, 255)
}
```

```javascript
/*本示例显示图表工作表 Chart1 中第一个图表组的高低点连线的透明度。*/
function test() {
    let chartGroup = Application.Charts.Item("Chart1").ChartObjects(1).Chart.ChartGroups(1)
    let chartFormat = chartGroup.HiLoLines.Format
    console.log(chartFormat.Line.Transparency)
}
```
