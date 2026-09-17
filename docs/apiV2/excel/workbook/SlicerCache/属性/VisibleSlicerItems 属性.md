---
title: "VisibleSlicerItems 属性"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/SlicerCache/%E5%B1%9E%E6%80%A7/VisibleSlicerItems%20%E5%B1%9E%E6%80%A7.html
---
# VisibleSlicerItems (属性)

返回一个 **[SlicerItems](./../../SlicerItems/SlicerItems 对象.html)** 集合，该集合包含指定的切片器缓存中所有可见项的集合。只读。

## 说明

**VisibleSlicerItems** 属性只适用于基于数据透视表，且该数据透视表又基于工作簿区域或列表的切片器 (**SlicerCache**.**SourceType** = **xlDatabase**)。对于连接到 OLAP 数据源的切片器 (**SlicerCache**.**OLAP** = **True**)，如果试图访问其 **VisibleSlicerItems** 属性，将会生成运行时错误。

## 返回值

SlicerItems

## 示例

```javascript
/*本示例显示活动工作簿上名为“切片器_score”的切片器缓存可见项的数量。*/
function test() {
    console.log(ActiveWorkbook.SlicerCaches("切片器_score").VisibleSlicerItems.Count)
}
```

```javascript
/*本示例显示第一张工作簿上第一个切片器缓存的可见项集合的第一项的名称。*/
function test() {
    let vsScItems = Workbooks.Item(1).SlicerCaches(1).VisibleSlicerItems
    console.log(vsScItems.Item(1).Name)
}
```
