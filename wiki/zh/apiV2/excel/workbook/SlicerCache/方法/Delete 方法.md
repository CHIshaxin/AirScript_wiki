---
title: "Delete (方法)"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/SlicerCache/%E6%96%B9%E6%B3%95/Delete%20%E6%96%B9%E6%B3%95.html
updated: 1789700916000
hash: D4DSa8JW
---

# Delete (方法)

删除指定的切片器缓存及其关联的切片器。

## 说明

若要脱离切片器缓存，而单独删除一个特定的切片器，那么应改用 **[Slicer](./../../Slicer/Slicer 对象.html)** 对象的 **Delete** 方法。

## 示例

```javascript
/*本示例删除活动工作簿“切片器_name”切片器缓存，以及与该切片器缓存关联的“name”切片器。*/
function test() {
    ActiveWorkbook.SlicerCaches("切片器_name").Delete()
}
```

```javascript
/*本示例判断如果第一张工作簿中第二个切片器缓存的名称为“切片器_date”，则删除该切片器缓存及关联的切片器。*/
function test() {
    let slcCache = Workbooks.Item(1).SlicerCaches(2)
    if (slcCache.Name == "切片器_date") {
        slcCache.Delete()
    }
}
```
