---
title: "AddPivotTable 方法"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/SlicerPivotTables/%E6%96%B9%E6%B3%95/AddPivotTable%20%E6%96%B9%E6%B3%95.html
---
# AddPivotTable (方法)

将对数据透视表的引用添加到 **[SlicerPivotTables](./../../SlicerPivotTables/SlicerPivotTables 对象.html)** 集合。

## 说明

当数据透视表添加到 **[SlicerPivotTables](./../../SlicerPivotTables/SlicerPivotTables 对象.html)** 集合时，可按其父级 **SlicerCache** 及其关联的切片器进行筛选。

## 参数

| 属性 | 数据类型 | 必填 | 说明 |
| --- | --- | --- | --- |
| PivotTable | PivotTable | 必填 | 一个代表要添加的数据透视表的 **[PivotTable](./../../PivotTable/PivotTable 对象.html)** 对象。 |

## 返回值

null

## 示例

```javascript
/*本示例将工作表 Sheet1 上数据透视表 pivot2 添加到活动工作簿上名为“切片器_id”的切片器缓存中。*/
function test() {
    let pvtTables = ActiveWorkbook.SlicerCaches("切片器_id").PivotTables
    pvtTables.AddPivotTable(Worksheets.Item("Sheet1").PivotTables("pivot2"))
}
```

```javascript
/*本示例将工作表 Sheet1 上数据透视表 pivot3 添加到与活动工作表上数据透视表 pivot1 的第一个切片器关联的切片器缓存中。*/
function test() {
    let pvtTables = ActiveSheet.PivotTables("pivot1").Slicers.Item(1).SlicerCache.PivotTables
    pvtTables.AddPivotTable(Worksheets.Item("Sheet1").PivotTables("pivot3"))
}
```
