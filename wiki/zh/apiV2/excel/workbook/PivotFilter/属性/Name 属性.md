---
title: "Name (属性)"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/PivotFilter/%E5%B1%9E%E6%80%A7/Name%20%E5%B1%9E%E6%80%A7.html
updated: 1789700916000
hash: BaMWoVn1
---

# Name (属性)

此属性提供选项可指定引用筛选。索引值可能发生更改，因此不能依赖于这些值获得准确的引用。

## 示例

```javascript
/*本示例显示活动工作表中第一个数据透视表的字段“name”的第一个筛选器的名称。*/
function test() {
    let pvtFilter = ActiveSheet.PivotTables(1).PivotFields("name").PivotFilters.Item(1)
    console.log(pvtFilter.Name)
}
```

```javascript
/*本示例显示活动工作表中数据透视表的字段“id”的第二个筛选器的名称。*/
function test() {
    let pvtFilter = ActiveSheet.Range("I1").PivotTable.PivotFields("id").PivotFilters.Item(2)
    console.log(`筛选器的名称：${pvtFilter.Name}`)
}
```
