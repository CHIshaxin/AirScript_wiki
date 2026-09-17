---
title: "Color 属性"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/SparkColor/%E5%B1%9E%E6%80%A7/Color%20%E5%B1%9E%E6%80%A7.html
---
# Color (属性)

返回一个 **[FormatColor](./../../FormatColor/FormatColor 对象.html)** 对象，可使用该对象来设置迷你图中点的标记色。只读

## 返回值

**[FormatColor](./../../FormatColor/FormatColor 对象.html)**

## 示例

```javascript
/*本示例设置活动工作表上区域 A1:A4 中第一个迷你图组的水平轴的颜色。*/
function test() {
    let sparkColor = ActiveSheet.Range("A1:A4").SparklineGroups.Item(1).Axes.Horizontal.Axis
    sparkColor.Color.ColorIndex = 5
}
```

```javascript
/*本示例判断如果活动工作表上区域 A1:A4 中第一个迷你图组的迷你图的第一个数据点可见，则将该数据点设置为黑色。*/
function test() {
    let sparkColor = ActiveSheet.Range("A1:A4").SparklineGroups.Item(1).Points.Firstpoint
    if (sparkColor.Visible) {
        sparkColor.Color.ColorIndex = 1
    }
}
```
