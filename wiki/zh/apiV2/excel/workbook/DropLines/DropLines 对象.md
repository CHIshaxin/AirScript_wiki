---
title: "DropLines (对象)"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/DropLines/DropLines%20%E5%AF%B9%E8%B1%A1.html
updated: 1789700916000
hash: B_ENu_0Y
---

# DropLines (对象)

代表图表组中的垂直线。

## 说明

垂直线将图表中的数据点与 x 轴连接起来。只有折线图和面积图组可以有垂直线。此对象不是集合。没有代表单个垂直线的对象；或者打开图表组中所有数据点的垂直线，或者将其全部关闭。

如果 **HasDropLines** 属性为 **False**， **[DropLines](./../DropLines/DropLines 对象.html)** 对象的绝大部分属性将被禁用。

使用 **DropLines** 属性可返回 **[DropLines](./../DropLines/DropLines 对象.html)** 对象。

## 示例

```javascript
/*本示例打开嵌入的第一个图表的第一个图表组的垂直线，并将垂直线的颜色设置为红色。*/
function test() {
    Application.Worksheets.Item("Sheet1").ChartObjects(1).Activate()
    ActiveChart.ChartGroups(1).HasDropLines = true
    ActiveChart.ChartGroups(1).DropLines.Border.ColorIndex = 3
}
```

```javascript
/*本示例选中图表工作表 Chart1 中图表的第一个图表组的垂直线，并将该垂直线的阴影设置为不可见。*/
function test() {
    let droplines = Application.Charts.Item("Chart1").ChartObjects(1).Chart.ChartGroups(1).DropLines
    droplines.Select()
    Selection.Format.Shadow.Visible = false
}
```
