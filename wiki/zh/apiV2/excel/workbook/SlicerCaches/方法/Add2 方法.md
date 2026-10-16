---
title: "Add2 (方法)"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/SlicerCaches/%E6%96%B9%E6%B3%95/Add2%20%E6%96%B9%E6%B3%95.html
updated: 1789700916000
hash: DAjlQAIb
---

# Add2 (方法)

将新的 **[SlicerCache](./../../SlicerCache/SlicerCache 对象.html)** 对象添加到集合。

## 参数

| 属性 | 数据类型 | 必填 | 说明 |
| --- | --- | --- | --- |
| Source | any | 必填 | 新 **[SlicerCache](./../../SlicerCache/SlicerCache 对象.html)** 对象将基于的数据源。 传给 *Source* 参数的自变量可以是 **WorkbookConnection** 对象、 **[PivotTable](./../../PivotTable/PivotTable 对象.html)** 对象或字符串。如果传递 **数据透视表** 对象，则关联的 **PivotCache** 对象将用作数据源。 如果传递了字符串，则会将其解释为 **WorkbookConnection** 对象的名称，如果不存在此类 **WorkbookConnection** 对象，则会生成运行时错误。 |
| SourceField | any | 必填 | 数据源中要作为筛选依据的字段的名称。 对于非 OLAP 数据源，请使用切片器所基于的 **PivotCache** 对象的 **[PivotField](./../../PivotField/PivotField 对象.html)** 对象，或使用该对象的唯一名称 (**PivotField** 的值。**Name**属性) 。对于 OLAP 数据源，应使用 **SlicerCache** 所基于的层次结构的 MDX 唯一名称。 还可以指定 OLAP 层次结构的一个级别，ET 将使用对应的层次结构。 |
| Name | any | 可选 | ET 用于引用切片器缓存的名称 (**SlicerCache** 的值。**Name** 属性) 。 如果省略此参数，ET 将生成一个名称。默认情况下，ET 将“Slicer_”与 **PivotField** 的值连接在一起。具有非 OLAP 数据源或 **CubeField** 值的切片器的 **Caption** 属性。具有 OLAP 数据源的切片器的 **Caption** 属性 (将任何空格替换为“_”) 。）如果需要使该名称在工作簿命名空间中唯一，ET 会将一个整数添加到生成的名称的末尾。 如果指定工作簿命名空间中已经存在的名称，**Add2** 方法将失败。 |
| SlicerCacheType | any | 可选 | 指定切片器或切片器缓存的类型。 |

## 返回值

SlicerCache

## 示例

```javascript
/*本示例基于工作表 Sheet1 上数据透视表在第一张工作簿上添加新的切片器缓存，并显示该切片器缓存的名称。*/
function test() {
    let pvtTable = Worksheets.Item("Sheet1").Range("I1").PivotTable
    let sliCache = Workbooks.Item(1).SlicerCaches.Add2(pvtTable, "name", "name1")
    console.log(sliCache.Name)
}
```

```javascript
/*本示例在活动工作簿上添加新的切片器缓存，并显示该工作簿上切片器缓存的数量。*/
function test() {
    let pvtTable = ActiveCell.PivotCell.PivotTable
    ActiveWorkbook.SlicerCaches.Add2(pvtTable, pvtTable.PivotFields("id"), "Slicer_id", xlSlicer)
    console.log(ActiveWorkbook.SlicerCaches.Count)
}
```
