---
title: "SlicerItems (属性)"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/SlicerCache/%E5%B1%9E%E6%80%A7/SlicerItems%20%E5%B1%9E%E6%80%A7.html
updated: 1789700916000
hash: D_eIpOyr
---

# SlicerItems (属性)

返回一个 **[SlicerItems](./../../SlicerItems/SlicerItems 对象.html)** 集合，该集合包含切片器缓存中所有项的集合。只读

## 说明

**[SlicerCache](./../../SlicerCache/SlicerCache 对象.html)** 对象的 **SlicerItems** 属性只适用于基于数据透视表，且该数据透视表又基于工作簿区域或列表的切片器 (**SlicerCache**.**SourceType** = **xlDatabase**)，或基于数据透视表，且该数据透视表又基于关系数据源的切片器（**SlicerCache**.**SourceType** = **xlExternal** 且 **SlicerCache**.**OLAP** = **False**）。对于连接到外部 OLAP 数据源的切片器 (**SlicerCache**.**OLAP** = **True**)，如果试图访问其 **SlicerItems** 属性，将会生成运行时错误。对于 OLAP 数据源，请改用 **SlicerCacheLevel** 对象的 **SlicerItems** 属性。

## 返回值

SlicerItems

## 示例

```javascript
/*本示例显示活动工作簿第一个切片器缓存中项的数量。*/
function test() {
    console.log(ActiveWorkbook.SlicerCaches(1).SlicerItems.Count)
}
```

```javascript
/*本示例取消选择第一张工作簿上名为“切片器_name”的切片器缓存的第三项。*/
function test() {
    let slicerItem = Workbooks.Item(1).SlicerCaches("切片器_name").SlicerItems(3)
    slicerItem.Selected = false
}
```
