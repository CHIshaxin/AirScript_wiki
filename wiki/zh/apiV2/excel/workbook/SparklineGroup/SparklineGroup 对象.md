---
title: "SparklineGroup (对象)"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/SparklineGroup/SparklineGroup%20%E5%AF%B9%E8%B1%A1.html
updated: 1789700916000
hash: BN4cJxTB
---

# SparklineGroup (对象)

代表一组迷你图。

## 说明

**[SparklineGroup](./../SparklineGroup/SparklineGroup 对象.html)** 对象可包含多个迷你图以及该迷你图组的属性设置，例如颜色和轴设置。每个迷你图都由一个 **[Sparkline](./../Sparkline/Sparkline 对象.html)** 对象表示。

使用 **Modify** 方法可在迷你图组中添加或删除迷你图。使用 **ModifyLocation** 方法可更改迷你图的位置，使用 **ModifySourceData** 方法则可以更改源数据所在的区域。

## 示例

```javascript
/*本示例会在 A1:A4 处创建一组绑定到 Sheet1!B1:E4 区域中的源数据的列迷你图，并更改系列颜色以便用红色显示各个列。*/
function test() {
    let sparklineGroup = Range("A1:A4").SparklineGroups.Add(xlSparkColumn, "Sheet1!B1:E4")
    sparklineGroup.SeriesColor.Color = RGB(255, 0, 0)
}
```

```javascript
/*本示例显示活动工作表上单元格 H1 中第一个迷你图组中迷你图的高点是否可见。*/
function test() {
    let sparklineGroup = ActiveSheet.Range("H1").SparklineGroups.Item(1)
    console.log(sparklineGroup.Points.Highpoint.Visible)
}
```
