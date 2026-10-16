---
title: "Value1 属性"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/PivotFilter/%E5%B1%9E%E6%80%A7/Value1%20%E5%B1%9E%E6%80%A7.html
---
# Value1 (属性)

此属性是由用户提供的用于定义透视字段筛选器的参数。**Variant** 类型，只读。

## 示例

```javascript
/*本示例显示第一张工作表中数据透视表的字段“id”的第二个筛选器的参数。*/
function test() {
    let pvtField = Worksheets.Item(1).Range("I1").PivotTable.PivotFields("id")
    console.log(pvtField.PivotFilters.Item(2).Value1)
}
```

```javascript
/*本示例显示活动工作表中第一个数据透视表的“name”字段的第一个筛选器的参数。*/
function test() {
    let pvtFilter = ActiveSheet.PivotTables(1).PivotFields("name").PivotFilters.Item(1)
    console.log(pvtFilter.Value1)
}
```
