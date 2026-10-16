---
title: "PivotTables (属性)"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/SlicerCache/%E5%B1%9E%E6%80%A7/PivotTables%20%E5%B1%9E%E6%80%A7.html
updated: 1789700916000
hash: CKUOW4rj
---

# PivotTables (属性)

返回一个 **[SlicerPivotTables](./../../SlicerPivotTables/SlicerPivotTables 对象.html)** 集合，其中包含有关切片器缓存当前正在筛选的数据透视表的信息。只读

## 说明

如果与指定的 **SlicerCache** 关联的切片器未连接到任何数据透视表，那么 **PivotTables** 属性返回的 **[SlicerPivotTables](./../../SlicerPivotTables/SlicerPivotTables 对象.html)** 集合将为空。

## 返回值

PivotTables

## 示例

```javascript
/*本示例显示活动工作簿第一个切片器缓存正在筛选的第一张数据透视表的名称。*/
function test() {
    let pvtTables = ActiveWorkbook.SlicerCaches(1).PivotTables
    console.log(pvtTables.Item(1).Name)
}
```

```javascript
/*本示例显示第一张工作簿中名为“切片器_date”的切片器缓存正在筛选的第一张数据透视表的字段的数量。*/
function test() {
    let pvtTables = Workbooks.Item(1).SlicerCaches("切片器_date").PivotTables
    console.log(pvtTables.Item(1).PivotFields().Count)
}
```
