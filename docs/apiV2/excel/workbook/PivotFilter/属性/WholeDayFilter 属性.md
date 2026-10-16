---
title: "WholeDayFilter 属性"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/PivotFilter/%E5%B1%9E%E6%80%A7/WholeDayFilter%20%E5%B1%9E%E6%80%A7.html
---
# WholeDayFilter (属性)

设置或获取日期筛选器的筛选语义。 可读写 **Boolean** 类型。

## 说明

对于时间线，始终返回 **True**；将此项设置为 **False** 时返回运行时错误。

对于非日期筛选器，任何访问都返回运行时错误。

## 示例

```javascript
/*本示例显示第一张工作表中数据透视表的字段“date”的第一个筛选器是否对时间线实施筛选。*/
function test() {
    let pvtField = Worksheets.Item(1).Range("I1").PivotTable.PivotFields("date").PivotFilters.Item(1)
    console.log(pvtField.WholeDayFilter)
}
```

```javascript
/*本示例将活动工作表中数据透视表的字段“date”的第一个筛选器设置为对时间线实施筛选。*/
function test() {
    let pvtField = ActiveSheet.PivotTables(1).PivotFields("date").PivotFilters.Item(1)
    pvtField.WholeDayFilter = true
}
```
