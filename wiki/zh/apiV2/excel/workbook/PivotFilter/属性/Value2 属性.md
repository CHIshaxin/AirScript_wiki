---
title: "Value2 (属性)"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/PivotFilter/%E5%B1%9E%E6%80%A7/Value2%20%E5%B1%9E%E6%80%A7.html
updated: 1789700916000
hash: BRBpLhzk
---

# Value2 (属性)

此属性是由用户提供的用于定义透视字段筛选器的参数。**Variant** 类型，只读。

## 示例

```javascript
/*本示例显示第一张工作表中数据透视表的第一个字段的第一个筛选器的参数。*/
function test() {
    let pvtField = Worksheets.Item(1).Range("I1").PivotTable.PivotFields(1)
    console.log(pvtField.PivotFilters.Item(1).Value2)
}
```

```javascript
/*本示例显示活动工作表中第一个数据透视表的第三个字段的第二个筛选器的参数。*/
function test() {
    let pvtFilter = ActiveSheet.PivotTables(1).PivotFields(3).PivotFilters.Item(2)
    console.log(pvtFilter.Value2)
}
```
