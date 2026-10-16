---
title: "SeriesColor (属性)"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/SparklineGroup/%E5%B1%9E%E6%80%A7/SeriesColor%20%E5%B1%9E%E6%80%A7.html
updated: 1789700916000
hash: DbqO9agY
---

# SeriesColor (属性)

返回一个代表迷你图组的主要系列颜色的 **[FormatColor](./../../FormatColor/FormatColor 对象.html)** 对象。只读。

## 返回值

FormatColor

## 示例

```javascript
/*本示例将活动工作表上单元格 J1 中第一个迷你图组的主系列颜色设置为红色。*/
function test() {
    let sparklineGroup = ActiveSheet.Range("J1").SparklineGroups.Item(1)
    sparklineGroup.SeriesColor.ColorIndex = 3
}
```

```javascript
/*本示例设置活动工作表上单元格 J1 中第一个迷你图组的主系列颜色。*/
function test() {
    let sparklineGroup = ActiveSheet.Range("J1").SparklineGroups.Item(1)
    sparklineGroup.SeriesColor.Color = RGB(255, 255, 0)
}
```
