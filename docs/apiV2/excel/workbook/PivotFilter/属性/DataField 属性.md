---
title: "DataField 属性"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/PivotFilter/%E5%B1%9E%E6%80%A7/DataField%20%E5%B1%9E%E6%80%A7.html
---
# DataField (属性)

此属性仅适用于非 OLAP 数据透视表，可为值筛选提供作为筛选依据的**值**字段（值区域中的透视字段）。只读 **PivotField**。

## 示例

```javascript
/*本示例显示工作表 Sheet1 中数据透视表的字段“name”的第一个筛选器中作为筛选依据的值字段的名称。*/
function test() {
    let pvtFilter = Worksheets.Item("Sheet1").Range("I1").PivotTable.PivotFields("name").PivotFilters.Item(1)
    console.log(pvtFilter.DataField.Value)
}
```

```javascript
/*本示例设置活动工作表中第一个数据透视表的字段“id”的第二个筛选器中作为筛选依据的值字段的公式。*/
function test() {
    let pvtFilter = ActiveSheet.PivotTables(1).PivotFields("id").PivotFilters.Item(2)
    pvtFilter.DataField.Formula = "score*0.9"
}
```
