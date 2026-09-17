---
title: "RemovePivotTable 方法"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/SlicerPivotTables/%E6%96%B9%E6%B3%95/RemovePivotTable%20%E6%96%B9%E6%B3%95.html
---
# RemovePivotTable (方法)

从 **[SlicerPivotTables](./../../SlicerPivotTables/SlicerPivotTables 对象.html)** 集合中删除对数据透视表的引用。

## 说明

从 **[SlicerPivotTables](./../../SlicerPivotTables/SlicerPivotTables 对象.html)** 集合中删除数据透视表时，将不能再按其父级 **SlicerCache** 以及与之关联的切片器进行筛选。

## 参数

| 属性 | 数据类型 | 必填 | 说明 |
| --- | --- | --- | --- |
| PivotTable | any | 必填 | 表示要删除的数据透视表的 **[PivotTable](./../../PivotTable/PivotTable 对象.html)** 对象，或者数据透视表在集合中的名称或索引。 |

## 返回值

null

## 示例

```javascript
/*本示例从与活动工作簿上名为“切片器_name”的切片器关联的缓存中删除 pivot2。*/
function test() {
    let pvtTables = ActiveWorkbook.SlicerCaches("切片器_name").PivotTables
    pvtTables.RemovePivotTable("pivot2")
}
```

```javascript
/*本示例从与活动工作表上数据透视表的第一个切片器关联的缓存中删除 pivot3。*/
function test() {
    let sliCache = ActiveSheet.Range("I1").PivotTable.Slicers.Item(1).SlicerCache
    sliCache.PivotTables.RemovePivotTable(ActiveSheet.PivotTables("pivot3"))
}
```
