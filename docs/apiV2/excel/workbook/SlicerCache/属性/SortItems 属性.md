---
title: "SortItems 属性"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/SlicerCache/%E5%B1%9E%E6%80%A7/SortItems%20%E5%B1%9E%E6%80%A7.html
---
# SortItems (属性)

返回或设置是否对切片器中的项排序。可读/写

## 说明

此属性的默认设置为 **xlSlicerSortAscending**。

**[SlicerCache](./../../SlicerCache/SlicerCache 对象.html)** 对象的 **SortItems** 属性只适用于基于连接到工作簿区域或列表的数据透视表的切片器 (**SlicerCache**.**SourceType** = **xlDatabase**)。对于连接到 OLAP 数据源的切片器 (**SlicerCache**.**OLAP** = **True**)，如果试图访问其 **SortItems** 属性，那么将生成运行时错误。对于 OLAP 数据源，请改用 **SlicerCacheLevel** 对象的 **SortItems** 属性。

## 返回值

**[XlSlicerSort](./../../枚举/XlSlicerSort 枚举.html)**

## 示例

```javascript
/*本示例显示活动工作簿上名为“切片器_name”的切片器缓存中项的排序顺序是否为 xlSlicerSortAscending。*/
function test() {
    console.log(ActiveWorkbook.SlicerCaches("切片器_name").SortItems == xlSlicerSortAscending)
}
```

```javascript
/*本示例将第一张工作簿上名为“切片器_id”的切片器缓存中项的排序顺序设置为 xlSlicerSortDescending。*/
function test() {
    let slcCache = Workbooks.Item(1).SlicerCaches("切片器_id")
    slcCache.SortItems = xlSlicerSortDescending
}
```
