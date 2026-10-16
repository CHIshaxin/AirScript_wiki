---
title: "ClearAllFilters (方法)"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/PivotField/%E6%96%B9%E6%B3%95/ClearAllFilters%20%E6%96%B9%E6%B3%95.html
updated: 1789700916000
hash: SBQZAmvs
---

# ClearAllFilters (方法)

调用此方法将删除当前应用于透视字段的所有筛选。包括从透视字段的 **[PivotFilters](./../../PivotFilters/PivotFilters 对象.html)** 集合中删除所有筛选以及删除应用于透视字段的所有手动筛选。如果透视字段位于报表筛选区域中，则所选项将被设置为默认项。

## 示例

```javascript
/*本示例删除活动工作表中第一张数据透视表应用于字段“id”的所有筛选。*/
function test() {
    ActiveSheet.PivotTables(1).PivotFields("id").ClearAllFilters()
}
```

```javascript
/*本示例删除工作表 Sheet1 中数据透视表应用于字段“grades”的所有筛选。*/
function test() {
    let pvtfield = Worksheets.Item("Sheet1").Range("I1").PivotTable.PivotFields("grades")
    pvtfield.ClearAllFilters()
}
```
