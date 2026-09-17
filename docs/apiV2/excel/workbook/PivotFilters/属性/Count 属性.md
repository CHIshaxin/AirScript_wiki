---
title: "Count 属性"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/PivotFilters/%E5%B1%9E%E6%80%A7/Count%20%E5%B1%9E%E6%80%A7.html
---
# Count (属性)

返回 **[PivotFilters](./../../PivotFilters/PivotFilters 对象.html)** 集合中的项数。只读。

## 示例

```javascript
/*本示例显示活动工作表中第一张数据透视表的字段“score”的筛选器的数量。*/
function test() {
    console.log(ActiveSheet.PivotTables(1).PivotFields("score").PivotFilters.Count)
}
```

```javascript
/*本示例显示第一张工作表中数据透视表的字段“id”的筛选器的数量。*/
function test() {
    let pvtFilters = Worksheets.Item(1).Range("I1").PivotTable.PivotFields("id").PivotFilters
    console.log(`筛选器的数量：${pvtFilters.Count}`)
}
```
