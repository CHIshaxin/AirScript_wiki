---
title: "PivotField 属性"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/PivotFilter/%E5%B1%9E%E6%80%A7/PivotField%20%E5%B1%9E%E6%80%A7.html
---
# PivotField (属性)

指定应用筛选器的透视字段。只读。

## 示例

```javascript
/*本示例显示应用于工作表 Sheet1 中数据透视表的第二个字段的第二个筛选器的字段的名称。*/
function test() {
    let pvtFilter = Worksheets.Item("Sheet1").Range("I1").PivotTable.PivotFields(2).PivotFilters.Item(2)
    console.log(pvtFilter.PivotField.Value)
}
```

```javascript
/*本示例选中应用于活动工作表中第一个数据透视表的第一个字段的第一个筛选器的字段的数据透视表项。*/
function test() {
    let pvtFilter = ActiveSheet.PivotTables(1).PivotFields(1).PivotFilters.Item(1)
    pvtFilter.PivotField.DataRange.Select()
}
```
