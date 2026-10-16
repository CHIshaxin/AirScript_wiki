---
title: "SparklineGroups (属性)"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/Range/%E5%B1%9E%E6%80%A7/SparklineGroups%20%E5%B1%9E%E6%80%A7.html
updated: 1789700916000
hash: C6lEQ3Cz
---

# SparklineGroups (属性)

返回一个 **[SparklineGroups](./../../SparklineGroups/SparklineGroups 对象.html)** 对象，该对象代表指定区域中的一个现有迷你图组。只读

## 返回值

**SparklineGroups**

## 示例

```javascript
/*此示例向区域 A1:A4 中添加一个迷你图组。 该组中的迷你图为列迷你图并且绑定到区域 B1:E4 中的数据。*/
function test() {
    Range("A1:A4").SparklineGroups.Add(xlSparkColumn, "B1:E4")
}
```

```javascript
/*本示例在 E1:E4 单元格区域中添加一个迷你图组，图组为折线图并关联 B1:C4 单元格区域的数据，显示 E1 单元格中迷你图组的数量；清除 E1 单元格中的迷你图组，再显示 E1 单元格中的迷你图组的数量*/
function test() {
    Range("E1:E4").SparklineGroups.Add(xlSparkLine, "B1:C4")
    console.log(Range("E1").SparklineGroups.Count)
    Range("E1").SparklineGroups.Clear()
    console.log(Range("E1").SparklineGroups.Count)
}
```
