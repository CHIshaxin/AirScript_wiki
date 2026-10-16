---
title: "ClearAllFilters (方法)"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/SlicerCache/%E6%96%B9%E6%B3%95/ClearAllFilters%20%E6%96%B9%E6%B3%95.html
updated: 1789700916000
hash: q9CBEw7v
---

# ClearAllFilters (方法)

清除切片器或时间线的筛选器，具体取决于切片器缓存类型。

## 示例

```javascript
/*本示例清除活动工作簿上所有切片器缓存的筛选器。*/
function test() {
    for (let i = 1; i <= ActiveWorkbook.SlicerCaches.Count; i++) {
        ActiveWorkbook.SlicerCaches(i).ClearAllFilters()
    }
}
```

```javascript
/*本示例清除第一张工作簿中名为“切片器_id”的切片器缓存的筛选器。*/
function test() {
    Workbooks.Item(1).SlicerCaches("切片器_id").ClearAllFilters()
}
```
