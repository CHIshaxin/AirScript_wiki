---
title: "PivotFilters 属性"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/PivotField/%E5%B1%9E%E6%80%A7/PivotFilters%20%E5%B1%9E%E6%80%A7.html
---
# PivotFilters (属性)

返回或设置指定 **[PivotField](./../../PivotField/PivotField 对象.html)** 对象的 PivotFilters。只读。

## 示例

```javascript
/*本示例显示活动工作表中第一张数据透视表的字段“score”的筛选器的数量。*/
function test() {
    console.log(ActiveSheet.PivotTables(1).PivotFields("score").PivotFilters.Count)
}
```

```javascript
/*本示例显示第一张工作表中数据透视表的字段“id”第一个筛选器的参数。*/
function test() {
    let pvtField = Worksheets.Item(1).Range("I3").PivotTable.PivotFields("id")
    console.log(pvtField.PivotFilters.Item(1).Value1)
}
```
