---
title: "Databar (对象)"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/Databar/Databar%20%E5%AF%B9%E8%B1%A1.html
updated: 1789700916000
hash: CEXFTb_9
---

# Databar (对象)

代表数据条条件格式规则。通过对范围应用数据条，有助于查看相对于其他单元格的单元格的值。

## 说明

所有条件格式对象都包含在 **[FormatConditions](./../FormatConditions/FormatConditions 对象.html)** 集合对象中，该集合对象是 **[Range](./../Range/Range 对象.html)** 集合的子项。可以使用 **[FormatConditions](./../FormatConditions/FormatConditions 对象.html)** 集合的 **Add** 或 **AddDatabar** 方法创建数据条格式设置规则。

可以使用 **[Databar](./../Databar/Databar 对象.html)** 对象的 **MinPoint** 和 **MaxPoint** 属性设置数据范围的最短和最长数据条的值。这些属性会返回 **[ConditionValue](./../ConditionValue/ConditionValue 对象.html)** 对象，使用该对象可指定如何计算阈值。

**[Databar](./../Databar/Databar 对象.html)** 对象还提供了另外一些属性，使用这些属性可指定在存在负值时显示的轴线，以及指定数据条的颜色和格式设置。

## 示例

```javascript
/*本示例将在活动工作表创建一个数据范围，然后对该范围应用数据条。为了明确显示中间值，使用ConditionValue对象将阈值的计算方式更改为百分点。*/
function test() {
    // Create a range of data with a couple of extreme values
    ActiveSheet.Range("D1").Value2 = 1
    ActiveSheet.Range("D2").Value2 = 45
    ActiveSheet.Range("D3").Value2 = 50
    ActiveSheet.Range("D2:D3").AutoFill(Range("D2:D8"))
    ActiveSheet.Range("D9").Value2 = 500
    Range("D1:D9").Select()

    // Create a data bar with default behavior
    let dataBar = Selection.FormatConditions.AddDatabar()
    console.log("Because of the extreme values, middle data bars are very similar")

    // The MinPoint and MaxPoint properties return a ConditionValue object
    // which you can use to change threshold parameters
    dataBar.MinPoint.Modify(xlConditionValuePercentile, 5)
    dataBar.MaxPoint.Modify(xlConditionValuePercentile, 75)
}
```

```javascript
/*本示例将活动工作表上区域 A1:A10 中第一个条件格式（数据条）的填充色设置为渐变色，并设置该数据条的展示方向。*/
function test() {
    let databar = ActiveSheet.Range("A1:A10").FormatConditions.Item(1)
    databar.BarFillType = xlDataBarFillGradient
    databar.Direction = xlRTL
}
```
