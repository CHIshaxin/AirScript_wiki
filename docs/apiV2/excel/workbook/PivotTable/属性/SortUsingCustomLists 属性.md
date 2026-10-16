---
title: "SortUsingCustomLists 属性"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/PivotTable/%E5%B1%9E%E6%80%A7/SortUsingCustomLists%20%E5%B1%9E%E6%80%A7.html
---
# SortUsingCustomLists (属性)

**SortUsingCustomLists** 属性控制最初在初始化透视字段并按 PivotItems 标题对 PivotItems 进行排列时，以及稍后在用户应用排序时，是否将自定义列表用于对字段项进行排序。可读/写 **Boolean** 类型。

## 说明

将此属性设置为 **False** 可优化具有多个项的字段的性能，并且它还允许不需要基于自定义列表排序的用户避免基于自定义列表进行排序。

## 示例

```javascript
/*本示例将自定义列表用于对数据透视表字段项进行排序*/
function test() {
    Worksheets.Item(1).PivotTables(1).SortUsingCustomLists = true
}
```

```javascript
/*本示例显示数据透视表的 SortUsingCustomLists 属性的数据类型，并设置为 false*/
function test() {
    console.log(typeof ActiveSheet.PivotTables(1).SortUsingCustomLists)
    ActiveSheet.PivotTables(1).SortUsingCustomLists = false
}
```
