---
title: "SlicerCache 属性"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/Slicer/%E5%B1%9E%E6%80%A7/SlicerCache%20%E5%B1%9E%E6%80%A7.html
---
# SlicerCache (属性)

返回与切片器关联的 **[SlicerCache](./../../SlicerCache/SlicerCache 对象.html)** 对象。只读。

## 返回值

SlicerCache

## 示例

```javascript
/*本示例显示工作表 Sheet1 上第一张数据透视表的第三个切片器的缓存中可见项集合的第三项的名称。*/
function test() {
    let sliCache = Worksheets.Item("Sheet1").PivotTables(1).Slicers.Item(3).SlicerCache
    console.log(sliCache.VisibleSlicerItems.Item(3).Name)
}
```

```javascript
/*本示例显示活动工作表上数据透视表的第一个切片器的缓存中第五个切片器项的值。*/
function test() {
    let sliItem = ActiveSheet.Range("I1").PivotTable.Slicers.Item(1).SlicerCache.SlicerItems(5)
    console.log(`第五个切片器项的值：${sliItem.Value}`)
}
```
