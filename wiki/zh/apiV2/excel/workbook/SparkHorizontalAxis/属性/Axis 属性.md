---
title: "Axis (属性)"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/SparkHorizontalAxis/%E5%B1%9E%E6%80%A7/Axis%20%E5%B1%9E%E6%80%A7.html
updated: 1789700916000
hash: Csfdc0Wm
---

# Axis (属性)

返回一个 **[SparkColor](./../../SparkColor/SparkColor 对象.html)** 对象，该对象指定迷你图水平轴的颜色。只读。

## 示例

```javascript
/*本示例设置活动工作表上单元格 H1 中第一个迷你图组的水平轴的颜色。*/
function test() {
    let shAxis = ActiveSheet.Range("H1").SparklineGroups(1).Axes.Horizontal
    shAxis.Axis.Color.ColorIndex = 3
}
```

```javascript
/*本示例将活动工作表上单元格 H1 中第一个迷你图组的水平轴的设置为不可见。*/
function test() {
    let shAxis = ActiveSheet.Range("H1").SparklineGroups(1).Axes.Horizontal
    shAxis.Axis.Visible = false
}
```
