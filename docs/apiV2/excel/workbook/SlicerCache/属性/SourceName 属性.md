---
title: "SourceName 属性"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/SlicerCache/%E5%B1%9E%E6%80%A7/SourceName%20%E5%B1%9E%E6%80%A7.html
---
# SourceName (属性)

返回切片器所连接到的数据源的名称。只读

## 说明

对于基于工作簿数据的切片器 (**SlicerCache**.**SourceType** = **xlDatabase**) 或基于非 OLAP 外部数据的切片器（**SlicerCache**.**SourceType** = **xlExternal** 且 **SlicerCache**.**OLAP** = **False**），则返回源数据中相应列的名称。

对于 OLAP 切片器 (**SlicerCache**.**OLAP** = **True**)，则返回切片器所基于的层次结构的 MDX 唯一名称。

## 返回值

String

## 示例

```javascript
/*本示例显示活动工作簿上名为“切片器_id”的切片器缓存所连接到的数据源的名称。*/
function test() {
    console.log(ActiveWorkbook.SlicerCaches("切片器_id").SourceName)
}
```

```javascript
/*本示例显示第一张工作簿上第一个切片器缓存所连接到的数据源的名称。*/
function test() {
    let slcCache = Workbooks.Item(1).SlicerCaches(1)
    console.log(`数据源名称：${slcCache.SourceName}`)
}
```
