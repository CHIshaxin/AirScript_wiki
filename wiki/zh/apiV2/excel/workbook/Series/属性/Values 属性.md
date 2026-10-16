---
title: "Values (属性)"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/Series/%E5%B1%9E%E6%80%A7/Values%20%E5%B1%9E%E6%80%A7.html
updated: 1789700916000
hash: Sw9LKdSH
---

# Values (属性)

返回或设置一个 **Variant** 值，它代表系列中所有值的集合。

## 说明

此属性值可以是工作表中的某一区域或常量值的数组，但不能是两者的组合。有关详细信息，请参阅此示例。

## 示例

```javascript
/*此示例设置区域的系列值。*/
function test() {
    let series = Application.Charts.Item("Chart1").ChartObjects(1).Chart.SeriesCollection(1)
    series.Values = Application.Worksheets.Item("Sheet1").Range("C5:T5")
}
```

```javascript
/*若要为每个单独的数据点赋以常量值，则必须使用数组。*/
function test() {
    let series = Application.Charts.Item("Chart1").ChartObjects(1).Chart.SeriesCollection(1)
    series.Values = [1, 3, 5, 7, 11, 13, 17, 19]
}
```
