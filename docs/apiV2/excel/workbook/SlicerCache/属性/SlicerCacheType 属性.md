---
title: "SlicerCacheType 属性"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/SlicerCache/%E5%B1%9E%E6%80%A7/SlicerCacheType%20%E5%B1%9E%E6%80%A7.html
---
# SlicerCacheType (属性)

返回切片器缓存的类型：切片器 (xlSlicer) 或时间线 (xlTimeline) 。只读。

## 返回值

XlSlicerCacheType

## 示例

```javascript
/*本示例显示活动工作簿第二个切片器缓存的类型是否为 xlSlicer。*/
function test() {
    console.log(ActiveWorkbook.SlicerCaches(2).SlicerCacheType == xlSlicer)
}
```

```javascript
/*本示例判断第一张工作簿中名为“切片器_id”的切片器缓存的类型，并通知用户。*/
function test() {
    let slcCache = Workbooks.Item(1).SlicerCaches("切片器_id")
    switch (slcCache.SlicerCacheType) {
        case xlSlicer:
            console.log("切片器缓存类型为切片器")
            break
        case xlTimeline:
            console.log("切片器缓存类型为时间线")
            break
    }
}
```
