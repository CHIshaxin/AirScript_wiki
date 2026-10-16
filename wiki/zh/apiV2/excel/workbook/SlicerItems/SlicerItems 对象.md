---
title: "SlicerItems (对象)"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/SlicerItems/SlicerItems%20%E5%AF%B9%E8%B1%A1.html
updated: 1789700916000
hash: iBmcM918
---

# SlicerItems (对象)

表示包含在 **SlicerCache** 或 **SlicerCacheLevel** 对象中的 **[SlicerItem](./../SlicerItem/SlicerItem 对象.html)** 对象的集合。

## 说明

如果切片器基于工作簿中的数据或非 OLAP 外部数据，那么若要访问表示该切片器中项目的 **[SlicerItems](./../SlicerItems/SlicerItems 对象.html)** 集合，请使用与该切片器关联的 **[SlicerCache](./../SlicerCache/SlicerCache 对象.html)** 对象的 **SlicerItems** 属性。

如果切片器基于 OLAP 数据连接，那么若要访问表示该切片器中项目的 **[SlicerItems](./../SlicerItems/SlicerItems 对象.html)** 集合，请使用代表层次结构级别的 **SlicerCacheLevel** 对象的 **SlicerItems** 属性。

## 示例

```javascript
/*本示例显示活动工作表上数据透视表第一个切片器的缓存中切片器项的数量。*/
function test() {
    let sliItems = ActiveSheet.Range("I1").PivotTable.Slicers.Item(1).SlicerCache.SlicerItems
    console.log(`切片器项的数量：${sliItems.Count}`)
}
```

```javascript
/*本示例将活动工作簿上名为“切片器_score”的切片器缓存中切片器项的名称逐个赋值到单元格中。*/
function test() {
    let sliItems = ActiveWorkbook.SlicerCaches("切片器_score").SlicerItems
    for (let i = 1; i <= sliItems.Count; i++) {
        Cells.Item(11, i).Value2 = sliItems.Item(i).Name
    }
}
```
