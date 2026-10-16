---
title: "IsMemberPropertyFilter 属性"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/PivotFilter/%E5%B1%9E%E6%80%A7/IsMemberPropertyFilter%20%E5%B1%9E%E6%80%A7.html
---
# IsMemberPropertyFilter (属性)

指定标签筛选器是基于字段的成员属性的 PivotItem 标题还是基于透视字段本身的 PivotItem 标题。只读 **Boolean** 类型。

## 说明

此属性的默认值为 **False**。

如果标签筛选器基于透视字段的成员属性的 PivotItem 标题，则返回 **True**；或者，如果筛选器基于透视字段的 PivotItem 标题，则返回 **False**。此属性只对标签筛选器和至少有一个成员属性的 OLAP 透视字段有效。

## 示例

```javascript
/*本示例显示工作表 Sheet1 中数据透视表的字段“name”的第一个筛选器是否基于透视字段本身的 PivotItem 标题。*/
function test() {
    let pvtFilter = Worksheets.Item("Sheet1").Range("I1").PivotTable.PivotFields("name").PivotFilters.Item(1)
    console.log(pvtFilter.IsMemberPropertyFilter)
}
```

```javascript
/*本示例显示活动工作表中第一个数据透视表的字段“score”的第一个筛选器是基于字段的成员属性的 PivotItem 标题还是基于透视字段本身的 PivotItem 标题，并通知用户。*/
function test() {
    let pvtFilter = ActiveSheet.PivotTables(1).PivotFields("score").PivotFilters.Item(1)
    if (pvtFilter.IsMemberPropertyFilter) {
        console.log("标签筛选器是基于透视字段本身的 PivotItem 标题")
    } else {
        console.log("标签筛选器是基于字段的成员属性的 PivotItem 标题")
    }
}
```
