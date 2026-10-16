---
title: "ShowAllItems (属性)"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/SlicerCache/%E5%B1%9E%E6%80%A7/ShowAllItems%20%E5%B1%9E%E6%80%A7.html
updated: 1789700916000
hash: BoxKFcYn
---

# ShowAllItems (属性)

返回或设置连接到指定切片器缓存的切片器是否显示已从对应的透视缓存中删除的项。可读/写。

## 说明

当 **ShowAllItems** 属性设置为 **True**（默认）时，已从源数据中删除的项将在连接到指定切片器缓存的切片器中显示。**ShowAllItems** 属性对应于 **“切片器设置”** 对话框中的 **“显示从数据源删除的项目”** 复选框的设置。

**ShowAllItems** 属性只适用于基于工作簿区域或列表的切片器 (**SlicerCache**.**SourceType** = **xlDatabase**) 或基于关系数据源的切片器（**SlicerCache**.**SourceType** = **xlExternal** 且 **SlicerCache**.**OLAP** = **False**）。如果切片器连接到基于外部 OLAP 数据源的数据透视表 (**SlicerCache**.**OLAP** = **True**)，那么尝试对该切片器设置 **ShowAllItems** 属性将生成运行时错误。

## 返回值

Boolean

## 示例

```javascript
/*本示例显示活动工作簿第二个切片器缓存关联的切片器是否显示已从对应的透视缓存中删除的项。*/
function test() {
    console.log(ActiveWorkbook.SlicerCaches(2).ShowAllItems)
}
```

```javascript
/*本示例将第一张工作簿中名为“切片器_id”的切片器缓存关联的切片器设置为不显示已从对应的透视缓存中删除的项。*/
function test() {
    let slcCache = Workbooks.Item(1).SlicerCaches("切片器_id")
    slcCache.ShowAllItems = false
}
```
