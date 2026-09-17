---
title: "Add 方法"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/Slicers/%E6%96%B9%E6%B3%95/Add%20%E6%96%B9%E6%B3%95.html
---
# Add (方法)

创建新的切片器，并返回 **[Slicer](./../../Slicer/Slicer 对象.html)** 对象。

## 参数

| 属性 | 数据类型 | 必填 | 说明 |
| --- | --- | --- | --- |
| SlicerDestination | any | 必填 | 一个指定工作表名称的 **String**，或是一个代表工作表（将在其中放置生成的切片器）的 **[Worksheet](./../../Worksheet/Worksheet 对象.html)** 对象。目标工作表必须位于表达式指定的 **[Slicers](./../../Slicers/Slicers 对象.html)** 对象所在的工作簿中。 |
| Level | any | 可选 | 如果是 OLAP 数据源，则为创建切片器所基于的级别的序号或多维表达式 (MDX) 名称。非 OLAP 数据源不支持此参数。 |
| Name | any | 可选 | 切片器的名称。如果未指定，ET 会自动生成一个名称。该名称必须在工作簿中的所有切片器中是唯一的。 |
| Caption | any | 可选 | 切片器的标题。 |
| Top | any | 可选 | 切片器相对于工作表上单元格 A1 左上角的初始垂直位置（以磅为单位）。 |
| Left | any | 可选 | 切片器相对于工作表上单元格 A1 左上角的初始水平位置（以磅为单位）。 |
| Width | any | 可选 | 切片器控件的初始宽度（以磅为单位）。 |
| Height | any | 可选 | 切片器控件的初始高度（以磅为单位）。 |

## 返回值

Slicer

## 示例

```javascript
/*本示例在活动工作簿上添加新的切片器缓存，并基于该切片器缓存创建新的切片器。*/
function test() {
    let pvtTable = ActiveCell.PivotCell.PivotTable
    let sliCache = ActiveWorkbook.SlicerCaches.Add2(pvtTable, pvtTable.PivotFields("id"), "Slicer_id", xlSlicer)
    sliCache.Slicers.Add(ActiveSheet, null, "Nameid", "nameid", 252, 522, 144, 216)
}
```

```javascript
/*本示例基于工作表 Sheet1 上数据透视表的第一个切片器的缓存在工作表 Sheet2 上创建新的切片器。*/
function test() {
    let sliCache = Worksheets.Item("Sheet1").Range("I1").PivotTable.Slicers.Item(1).SlicerCache
    sliCache.Slicers.Add(Worksheets.Item("Sheet2"), null, "Name1", "name1", 0, 0, 150, 200)
}
```
