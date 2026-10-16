---
title: "SparkPoints 对象"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/SparkPoints/SparkPoints%20%E5%AF%B9%E8%B1%A1.html
---
# SparkPoints (对象)

表示迷你图上数据点的标记设置。

## 说明

使用 **[SparkPoints](./../SparkPoints/SparkPoints 对象.html)** 对象可设置迷你图上数据点标记的颜色和可见性。使用 **[SparklineGroup](./../SparklineGroup/SparklineGroup 对象.html)** 对象的 **Points** 属性可返回 **[SparkPoints](./../SparkPoints/SparkPoints 对象.html)** 对象。 **[SparkPoints](./../SparkPoints/SparkPoints 对象.html)** 对象的属性对应于 **“显示”** 部分中 **“高点”**、**“低点”**、**“负点”**、**“首点”**、**“尾点”** 和 **“标记”** 复选框的设置，以及功能区的 **“迷你图工具设计”** 选项卡上 **“样式”** 部分中的 **“标记颜色”** 下拉列表内各项的设置。

## 示例

```javascript
/*本示例将活动工作表上区域 A1:A4 中第一个迷你图组的迷你图的最高和最低数据点设置为可见，并设置最高数据点的颜色。*/
function test() {
    let points = ActiveSheet.Range("A1:A4").SparklineGroups.Item(1).Points
    points.Highpoint.Visible = true
    points.Lowpoint.Visible = true
    points.Highpoint.Color.ColorIndex = 4
}
```

```javascript
/*本示例将活动工作表上区域 A1:A4 中第一个迷你图组的迷你图的最后一个数据点设置为可见，并设置其颜色。*/
function test() {
    let sparkColor = ActiveSheet.Range("A1:A4").SparklineGroups.Item(1).Points.Lastpoint
    sparkColor.Visible = true
    sparkColor.Color.Color = RGB(255, 0, 0)
}
```
