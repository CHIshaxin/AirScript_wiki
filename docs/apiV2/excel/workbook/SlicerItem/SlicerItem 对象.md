---
title: "SlicerItem 对象"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/SlicerItem/SlicerItem%20%E5%AF%B9%E8%B1%A1.html
---
# SlicerItem (对象)

表示切片器中的一个项目。

## 说明

若要访问表示切片器中当前选中的按钮的 **[SlicerItem](./../SlicerItem/SlicerItem 对象.html)** 对象，请使用 **[Slicer](./../Slicer/Slicer 对象.html)** 对象的 **ActiveItem** 属性。如果切片器筛选数据透视表，那么若要访问表示该切片器中所有项目的 **[SlicerItems](./../SlicerItems/SlicerItems 对象.html)** 集合，请使用与该 **[Slicer](./../Slicer/Slicer 对象.html)** 对象关联的 **[SlicerCache](./../SlicerCache/SlicerCache 对象.html)** 对象的 **SlicerItems** 属性。如果切片器筛选 OLAP 层次结构级别，那么若要访问表示该切片器中项目的 **[SlicerItems](./../SlicerItems/SlicerItems 对象.html)** 集合，请使用表示该层次结构级别的 **SlicerCacheLevel** 对象的 **SlicerItems** 属性。

## 示例

```javascript
/*本示例选中第一张工作簿上名为“切片器_id”的切片器缓存中所有切片器项。*/
function test() {
    let sliItems = Workbooks.Item(1).SlicerCaches("切片器_id").SlicerItems
    for (let i = 1; i <= sliItems.Count; i++) {
        sliItems.Item(i).Selected = true
    }
}
```

```javascript
/*本示例显示第一张工作簿上名为“切片器_name”的切片器缓存中第二个切片器项的源名称。*/
function test() {
    let sliItem = Workbooks.Item(1).SlicerCaches("切片器_name").SlicerItems(2)
    console.log(sliItem.SourceName)
}
```
