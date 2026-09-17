---
title: "UpBars 对象"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/UpBars/UpBars%20%E5%AF%B9%E8%B1%A1.html
---
# UpBars (对象)

涨柱线将图表组中第一个系列的数据点与最后一个系列中相应的有较大值的数据点连接起来（从第一个系列向上生长）。只有至少包含两个系列的二维折线图才能有涨柱线。此对象不是集合。没有代表单个涨柱线的对象；或者打开图表组中所有数据点的涨柱线，或者将其全部关闭。

## 说明

使用 **UpBars** 属性可返回 **[UpBars](./../UpBars/UpBars 对象.html)** 对象。

## 示例

```javascript
/*本示例打开工作表“Sheet5”上嵌入的第一个图表中第一个图表组的涨跌柱线，然后将涨柱线的颜色设置为蓝色，而将跌柱线设置为红色。*/
function test() {
    let chartgroup = Application.Worksheets.Item("Sheet5").ChartObjects(1).Chart.ChartGroups(1)
    chartgroup.HasUpDownBars = true
    chartgroup.UpBars.Interior.Color = RGB(0, 0, 255)
    chartgroup.DownBars.Interior.Color = RGB(255, 0, 0)
}
```

```javascript
/*本示例判断如果图表工作表 Chart1 中图表的第一个图表组的涨柱线的名称为“涨柱线 1”，则删除涨跌柱线。*/
function test() {
    let upbars = Application.Charts.Item("Chart1").ChartObjects(1).Chart.ChartGroups(1).UpBars
    if (upbars.Name == "涨柱线 1") {
        upbars.Delete()
    }
}
```
