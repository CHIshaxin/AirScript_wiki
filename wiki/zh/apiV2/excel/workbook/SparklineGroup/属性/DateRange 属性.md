---
title: "DateRange (属性)"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/SparklineGroup/%E5%B1%9E%E6%80%A7/DateRange%20%E5%B1%9E%E6%80%A7.html
updated: 1789700916000
hash: CnOMf9_X
---

# DateRange (属性)

获取或设置迷你图组的日期范围。可读/写。

## 说明

若要清除日期范围，请将此属性设置为空字符串。

日期范围必须是连续的一维范围。

日期范围可以在不同于 **Location** 和 **SourceData** 属性的工作表上。

日期范围中的空单元格和非日期值不显示。

## 返回值

String

## 示例

```javascript
/*本示例显示单元格 I1 中第一个迷你图组的日期范围。*/
function test() {
    let sparklineGroup = Range("I1").SparklineGroups.Item(1)
    console.log(sparklineGroup.DateRange)
}
```

```javascript
/*本示例设置活动工作表上单元格 I1 中第一个迷你图组的日期范围。*/
function test() {
    let sparklineGroup = ActiveSheet.Range("I1").SparklineGroups.Item(1)
    sparklineGroup.DateRange = "A1:F1"
}
```
