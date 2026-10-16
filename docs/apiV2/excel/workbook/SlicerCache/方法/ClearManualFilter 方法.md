---
title: "ClearManualFilter 方法"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/SlicerCache/%E6%96%B9%E6%B3%95/ClearManualFilter%20%E6%96%B9%E6%B3%95.html
---
# ClearManualFilter (方法)

清除与指定的切片器缓存关联的所有切片器的筛选。

## 说明

如果使用 **ClearManualFilter** 方法从一个切片器中清除筛选，但是有一个或多个切片器正交叉筛选该切片器中的数据，那么第一个切片器中的某些项将保持为淡灰。例如，如果有交叉筛选的 Customer 和 Country 切片器，并且您执行 `ActiveWorkbook.SlicerCaches("Slicer_Customer").ClearManualFilter()`，那么 Customer 切片器中的所有项要在您执行 `ActiveWorkbook.SlicerCaches("Slicer_Country").ClearManualFilter()` 后才会变得可用。有关交叉筛选的详细信息，请参阅 **CrossFilterType** 属性。

## 示例

```javascript
/*本示例清除活动工作簿上所有切片器缓存的筛选器。*/
function test() {
    for (let i = 1; i <= ActiveWorkbook.SlicerCaches.Count; i++) {
        ActiveWorkbook.SlicerCaches(i).ClearManualFilter()
    }
}
```

```javascript
/*本示例清除第一张工作簿中第三个切片器缓存的筛选器。*/
function test() {
    Workbooks.Item(1).SlicerCaches(3).ClearManualFilter()
}
```
