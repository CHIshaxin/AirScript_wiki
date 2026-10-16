---
title: "SparklineGroups (对象)"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/SparklineGroups/SparklineGroups%20%E5%AF%B9%E8%B1%A1.html
updated: 1789700916000
hash: BQin27zN
---

# SparklineGroups (对象)

代表迷你图组的集合。

## 说明

**[SparklineGroups](./../SparklineGroups/SparklineGroups 对象.html)** 对象可包含多个 **[SparklineGroup](./../SparklineGroup/SparklineGroup 对象.html)** 对象。

使用 **[Range](./../Range/Range 对象.html)** 对象的 **SparklineGroups** 属性可从现有 **[SparklineGroups](./../SparklineGroups/SparklineGroups 对象.html)** 集合的父区域中返回该集合。

使用 **Add** 方法可创建一组新迷你图。

使用 **Group** 方法可创建一组现有迷你图。

## 示例

```javascript
/*此示例选择区域 A1:A4 并组合该区域中的迷你图，然后将该区域中第一个迷你图组的正数据点设置为可见，并将其颜色设置为红色。*/
function test() {
    Range("A1:A4").Select()
    Selection.SparklineGroups.Group(Range("A1"))
    Selection.SparklineGroups.Item(1).Points.Markers.Visible = true
    Selection.SparklineGroups.Item(1).Points.Markers.Color.Color = RGB(255, 0, 0)
}
```

```javascript
/*此示例取消组合活动工作表上区域 A1:A4 上迷你图组，并删除 A1 单元格中迷你图组。*/
function test() {
    let sparklineGroups = ActiveSheet.Range("A1:A4").SparklineGroups
    sparklineGroups.Ungroup()
    Range("A1").SparklineGroups.ClearGroups()
}
```
