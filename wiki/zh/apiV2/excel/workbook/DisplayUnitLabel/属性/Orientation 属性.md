---
title: "Orientation (属性)"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/DisplayUnitLabel/%E5%B1%9E%E6%80%A7/Orientation%20%E5%B1%9E%E6%80%A7.html
updated: 1789700916000
hash: tupX6Lhz
---

# Orientation (属性)

返回或设置一个 **Variant** 值，它代表文本方向。

## 说明

此属性值可设为 -90 到 90 度之间的整数值或以下常量之一：

| **xlDownward** |
| --- |
| **xlHorizontal** |
| **xlUpward** |
| **xlVertical** |

## 示例

```javascript
/*本示例显示图表工作表 Chart1 中数值轴上的单位标签文本方向的度数。*/
function test() {
    let displayUnitLabel = Application.Charts.Item("Chart1").ChartObjects(1).Chart.Axes(xlValue).DisplayUnitLabel
    console.log(displayUnitLabel.Orientation)
}
```

```javascript
/*下例将 Sheet1 上的第一个图表数值轴上的单位标签文本方向设置成向下排列。*/
function test() {
    let displayUnitLabel = Application.Worksheets.Item("Sheet1").ChartObjects(1).Chart.Axes(xlValue).DisplayUnitLabel
    displayUnitLabel.Orientation = xlDownward
}
```
