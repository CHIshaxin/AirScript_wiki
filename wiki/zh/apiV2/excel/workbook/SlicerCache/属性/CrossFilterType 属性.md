---
title: "CrossFilterType (属性)"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/SlicerCache/%E5%B1%9E%E6%80%A7/CrossFilterType%20%E5%B1%9E%E6%80%A7.html
updated: 1789700916000
hash: Ca6iRw04
---

# CrossFilterType (属性)

返回或设置切片器是否与共享同一切片器缓存的其他切片器一起参与交叉筛选，以及如何显示交叉筛选。可读/写

## 说明

如果多个切片器与同一数据透视表关联，那么默认情况下，如果在一个切片器中作为筛选依据的项目在另一个切片器中没有对应的数据，那么这些项目将灰显。例如，如果有 Country 切片器和 State 切片器，您单击 Country 切片器中的某个国家/地区，那么不属于该国家/地区的所有州将灰显。此功能称为“交叉筛选”。

对应于 **CrossFilterType** 属性设置的用户界面设置是 **“切片器设置”** 对话框中的 **“直观地指示空数据项”** 和 **“最后显示空数据项”** 复选框。将 **CrossFilterType** 属性设置为 **xlSlicerCrossFilterShowItemsWithDataAtTop** 对应于同时选中 **“直观地指示空数据项”** 和 **“最后显示空数据项”** 复选框。将 **CrossFilterType** 属性设置为 **xlSlicerCrossFilterShowItemsWithNoData** 对应于只选中 **“直观地指示空数据项”** 复选框。同时清除这两个复选框对应于将 **CrossFilterType** 属性设置为 **xlSlicerNoCrossFilter**。

**[SlicerCache](./../../SlicerCache/SlicerCache 对象.html)** 对象的 **CrossFilterType** 属性不支持 OLAP 数据源 (**SlicerCache**.**OLAP** = **True**)。对于 OLAP 数据源，应改用 **SlicerCacheLevel** 对象的 **CrossFilterType** 属性。

## 返回值

**[XlSlicerCrossFilterType](./../../枚举/XlSlicerCrossFilterType 枚举.html)**

## 示例

```javascript
/*本示例显示活动工作簿上名为“切片器_name”的切片器缓存使用的交叉筛选的类型及其可视化方式是否为 xlSlicerCrossFilterShowItemsWithDataAtTop。*/
function test() {
    console.log(ActiveWorkbook.SlicerCaches("切片器_name").CrossFilterType == xlSlicerCrossFilterShowItemsWithDataAtTop)
}
```

```javascript
/*本示例将第一张工作簿上第一个切片器缓存使用的交叉筛选的类型及其可视化方式设置为 xlSlicerCrossFilterShowItemsWithNoData。*/
function test() {
    let sliCache = Workbooks.Item(1).SlicerCaches(1)
    sliCache.CrossFilterType = xlSlicerCrossFilterShowItemsWithNoData
}
```
