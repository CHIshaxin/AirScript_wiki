---
title: "SlicerCaches (对象)"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/SlicerCaches/SlicerCaches%20%E5%AF%B9%E8%B1%A1.html
updated: 1789700916000
hash: BMz1Oc7J
---

# SlicerCaches (对象)

表示与指定工作簿关联的切片器缓存的集合。

## 说明

使用 **[SlicerCaches](./../SlicerCaches/SlicerCaches 对象.html)** 集合的 **Item** 属性可返回与指定的 **[Workbook](./../Workbook/Workbook 对象.html)** 对象关联的 **[SlicerCache](./../SlicerCache/SlicerCache 对象.html)** 对象。可以使用 **Index** 属性的值，或者指定对象的 **Name** 属性来检索 **[SlicerCache](./../SlicerCache/SlicerCache 对象.html)** 对象。

## 示例

```javascript
/*本示例显示活动工作簿上名为“切片器_id”的切片器缓存所连接到的数据源的名称。*/
function test() {
    console.log(ActiveWorkbook.SlicerCaches.Item("切片器_id").SourceName)
}
```

```javascript
/*本示例在活动工作簿上添加新的切片器缓存，并显示该切片器缓存的名称。*/
function test() {
    let sliCache =  ActiveWorkbook.SlicerCaches.Add2(ActiveSheet.Range("I1").PivotTable, "date", "Slicer_date", xlSlicer)
    console.log(sliCache.Name)
}
```
