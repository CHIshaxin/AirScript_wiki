---
title: "Description 属性"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/PivotFilter/%E5%B1%9E%E6%80%A7/Description%20%E5%B1%9E%E6%80%A7.html
---
# Description (属性)

为 **[PivotFilter](./../../PivotFilter/PivotFilter 对象.html)** 对象提供可选说明。只读 **String**。

## 说明

说明字符串的最大长度为 255 个字符。

## 示例

```javascript
/*本示例显示工作表 Sheet1 中数据透视表的字段“name”的第一个筛选器的可选说明。*/
function test() {
    let pvtFilter = Worksheets.Item("Sheet1").Range("I1").PivotTable.PivotFields("name").PivotFilters.Item(1)
    console.log(pvtFilter.Description)
}
```

```javascript
/*本示例显示活动工作表中第一张数据透视表的字段“id”的第二个筛选器的可选说明。*/
function test() {
    let pvtFilter = ActiveSheet.PivotTables(1).PivotFields("id").PivotFilters.Item(2)
    console.log(pvtFilter.Description)
}
```
