---
title: "SparkColor 对象"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/SparkColor/SparkColor%20%E5%AF%B9%E8%B1%A1.html
---
# SparkColor (对象)

表示迷你图中的点的标记色。

## 说明

**[SparkColor](./../SparkColor/SparkColor 对象.html)** 对象对应于功能区 **“迷你图工具设计”** 选项卡的 **“样式”** 部分中 **“标记颜色”** 下拉列表中提供的项的设置。使用 **[SparkPoints](./../SparkPoints/SparkPoints 对象.html)** 对象的对应属性可设置这些项的颜色。

## 示例

```javascript
/*本示例将活动工作表上区域 A1:A4 中第一个迷你图组的迷你图的最高数据点设置为可见。*/
function test() {
    let sparkColor = ActiveSheet.Range("A1:A4").SparklineGroups.Item(1).Points.Highpoint
    sparkColor.Visible = true
}
```

```javascript
/*本示例将活动工作表上区域 A1:A4 中第一个迷你图组的迷你图的最后一个数据点设置为可见，并设置其颜色。*/
function test() {
    let sparkColor = ActiveSheet.Range("A1:A4").SparklineGroups.Item(1).Points.Lastpoint
    sparkColor.Visible = true
    sparkColor.Color.Color = RGB(255, 0, 255)
}
```
