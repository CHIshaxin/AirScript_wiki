---
title: "Position (属性)"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/PivotField/%E5%B1%9E%E6%80%A7/Position%20%E5%B1%9E%E6%80%A7.html
updated: 1789700916000
hash: B6qxkM02
---

# Position (属性)

返回或设置一个 Variant 值，他代表数据透视表上所有字段的位置，包括属于所有行、列、页以及数据方向上的字段（第一字段、第二字段、第三字段等等）。

## 示例

```javascript
/*本示例显示活动工作表中第一张数据透视表的字段“name”在行上的位置。*/
function test() {
    let pvtField = ActiveSheet.PivotTables(1).PivotFields("name")
    console.log(pvtField.Position)
}
```

```javascript
/*本示例将第一张工作表中数据透视表的字段“id”在行上的位置设置为“1”。*/
function test() {
    let pvtField = Worksheets.Item(1).Range("I1").PivotTable.PivotFields("id")
    pvtField.Position = 1
}
```
