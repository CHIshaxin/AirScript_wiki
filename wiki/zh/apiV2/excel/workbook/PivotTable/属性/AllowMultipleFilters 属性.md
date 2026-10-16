---
title: "AllowMultipleFilters (属性)"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/PivotTable/%E5%B1%9E%E6%80%A7/AllowMultipleFilters%20%E5%B1%9E%E6%80%A7.html
updated: 1789700916000
hash: B_GFCE3-
---

# AllowMultipleFilters (属性)

设置或检索一个值，该值指示是否可同时将多个筛选应用于一个透视字段。可读写 **Boolean**。

## 说明

默认值为 **False**。

如果此属性设置为 **True**，可将多个筛选应用于一个透视字段。如果此属性设置为 **False**，在对已应用了筛选的透视字段应用筛选时，将删除现有筛选，然后应用新筛选。当数据透视表具有应用了多个筛选的字段时，如果将此属性设置为 **False**，则将静默删除数据透视表中的所有筛选而不显示任何警告。但是通过用户界面删除筛选时，将显示警告。

## 示例

```javascript
/*本示例将数据透视表 Pivot1 设置为可同时将多个筛选应用于一个透视字段*/
function test() {
    Workbooks.Item(1).PivotTables("Pivot1").AllowMultipleFilters = true
}
```

```javascript
/*本示例显示数据透视表 Pivot1 的 AllowMultipleFilters 的数据类型，并将其设置为 false。*/
function test() {
    console.log("AllowMultipleFilters属性的数据类型为" + typeof ActiveSheet.PivotTables("Pivot1").AllowMultipleFilters)
    ActiveSheet.PivotTables("Pivot1").AllowMultipleFilters = false
}
```
