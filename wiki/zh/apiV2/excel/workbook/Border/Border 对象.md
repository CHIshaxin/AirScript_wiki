---
title: "Border (对象)"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/Border/Border%20%E5%AF%B9%E8%B1%A1.html
updated: 1789700916000
hash: BR_81eEz
---

# Border (对象)

代表对象的边框。

## 说明

大多数具有边框的对象（除 **Range**和 **[Style](./../Style/Style 对象.html)** 对象外）都将边框作为单一实体处理，而不管边框有几个边。整个边框必须作为一个整体单位返回。例如，使用 **TrendLine**对象的 **Border** 属性可返回此类对象的 **[Border](./../Border/Border 对象.html)** 对象。

## 示例

```javascript
/*本示例更改活动图表中趋势线的类型和线型。*/
function test() {
    let trendline = Application.ActiveChart.SeriesCollection(1).Trendlines(1)
    trendline.Type = xlLinear
    trendline.Border.LineStyle = xlDash
}
```

**Range** 和 **[Style](./../Style/Style 对象.html)** 对象具有四个分立的边框：左边框、右边框、顶部边框和底部边框，这四个边框既可单独返回，也可作为一个组同时返回。使用 **Borders** 属性可返回 **[Borders](./../Borders/Borders 对象.html)** 集合，该集合包含所有四个边框，并将这些边框视为一个单位。

```javascript
/*本示例向第一张工作表上的单元格 A1 右边缘的边框添加双边框。*/
function test() {
    Application.Worksheets.Item(1).Range("A1").Borders.Item(xlEdgeRight).LineStyle = xlDouble
}
```

使用 **Borders**(*index*)（其中 *index* 指定边框）可返回单个 **[Border](./../Border/Border 对象.html)** 对象。

```javascript
/*本示例设置单元格区域 A1:G1 的底部边框的颜色。*/
function test() {
    Range("A1:G1").Borders.Item(xlEdgeBottom).Color = RGB(255, 0, 0)
}
```

```javascript
/*本示例将边框的 Weight 属性设置为 xlThick 会诱使 LineStyle 属性变为 xlSolid，尽管之前已将其设置为 xlDashDotDot 。*/
function test() {
    let border = Selection.Borders.Item(xlDiagonalDown)
    border.Color = RGB(255, 0, 0)
    console.log("border.LineStyle = " + border.LineStyle)  //border.style = 1
    console.log("Set border.style = xlDashDotDot")  //Set border.style = DashDotDot
    border.LineStyle  = xlDashDotDot
    console.log("border.LineStyle = " + border.LineStyle)  //border.style = 5
    console.log("Set border.weight = xlThick")  //Set border.weight = Thick
    border.weight = xlThick
    console.log("border.LineStyle = " + border.LineStyle)  //border.style = 1
}
```

*Index* 可为以下 XlBordersIndex 常量之一：**xlDiagonalDown**、**xlDiagonalUp**、**xlEdgeBottom**、**xlEdgeLeft**、**xlEdgeRight**、**xlEdgeTop**、**xlInsideHorizontal** 或 **xlInsideVertical**。
