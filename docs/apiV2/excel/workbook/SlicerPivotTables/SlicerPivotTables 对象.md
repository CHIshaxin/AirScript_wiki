---
title: "SlicerPivotTables 对象"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/SlicerPivotTables/SlicerPivotTables%20%E5%AF%B9%E8%B1%A1.html
---
# SlicerPivotTables (对象)

表示与指定的 **[SlicerCache](./../SlicerCache/SlicerCache 对象.html)** 对象关联的数据透视表集合相关的信息。

## 说明

**[SlicerPivotTables](./../SlicerPivotTables/SlicerPivotTables 对象.html)** 集合包含有关切片器缓存当前正在筛选的数据透视表的信息。它提供的属性可用来确定与该切片器关联的数据透视表的数量，可用来检索表示正被筛选的数据透视表的 **[PivotTable](./../PivotTable/PivotTable 对象.html)** 对象。它还提供了在 **[SlicerPivotTables](./../SlicerPivotTables/SlicerPivotTables 对象.html)** 集合中添加和删除数据透视表的方法。如果与指定的 **SlicerCache** 关联的切片器未连接到任何数据透视表，那么 **[SlicerPivotTables](./../SlicerPivotTables/SlicerPivotTables 对象.html)** 集合将为空。

使用 **[SlicerCache](./../SlicerCache/SlicerCache 对象.html)** 对象的 **PivotTables** 属性可返回与 **SlicerCache** 关联的 **[SlicerPivotTables](./../SlicerPivotTables/SlicerPivotTables 对象.html)** 集合，而 SlicerCache 又可能与一个或多个切片器关联。

## 示例

```javascript
/*本示例显示活动工作表上数据透视表的第一个切片器关联的缓存中第二个数据透视表的值。*/
function test() {
    let sliCache = ActiveSheet.Range("I1").PivotTable.Slicers.Item(1).SlicerCache
    console.log(sliCache.PivotTables.Item(2).Value)
}
```

```javascript
/*本示例显示活动工作簿上名为“切片器_name”的切片器缓存关联的数据透视表的数量。*/
function test() {
    let pvtTables = ActiveWorkbook.SlicerCaches("切片器_name").PivotTables
    console.log(`数据透视表的数量：${pvtTables.Count}`)
}
```
