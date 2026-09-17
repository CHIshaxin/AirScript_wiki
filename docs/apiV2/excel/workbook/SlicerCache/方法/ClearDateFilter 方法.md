---
title: "ClearDateFilter 方法"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/SlicerCache/%E6%96%B9%E6%B3%95/ClearDateFilter%20%E6%96%B9%E6%B3%95.html
---
# ClearDateFilter (方法)

清除时间线的筛选器 (日期筛选器) 。

## 示例

```javascript
/*本示例清除第一张工作簿中第二个切片器缓存的日期筛选器。*/
function test() {
    Workbooks.Item(1).SlicerCaches(2).ClearDateFilter()
}
```

```javascript
/*本示例清除活动工作簿中名为“切片器_id”切片器缓存的日期筛选器。*/
function test() {
    ActiveWorkbook.SlicerCaches("切片器_id").ClearDateFilter()
}
```
