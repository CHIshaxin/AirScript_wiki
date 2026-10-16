---
title: "SlicerCacheType (属性)"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/Slicer/%E5%B1%9E%E6%80%A7/SlicerCacheType%20%E5%B1%9E%E6%80%A7.html
updated: 1789700916000
hash: BdEJZeab
---

# SlicerCacheType (属性)

返回切片器缓存的类型：切片器 (xlSlicer) 或时间线 (xlTimeline) 。只读。

## 返回值

XlSlicerCacheType

## 示例

```javascript
/*本示例显示活动工作表上第一张数据透视表的第二个切片器的切片器缓存类型是否为 xlSlicer。*/
function test() {
    console.log(ActiveSheet.PivotTables(1).Slicers.Item(2).SlicerCacheType == xlSlicer)
}
```

```javascript
/*本示例判断第一张工作表上数据透视表的第一个切片器的切片器缓存的类型，并通知用户。*/
function test() {
    let slicer = Worksheets.Item(1).Range("I1").PivotTable.Slicers.Item(1)
    switch (slicer.SlicerCacheType) {
        case xlSlicer:
            console.log("切片器缓存类型为切片器")
            break
        case xlTimeline:
            console.log("切片器缓存类型为时间线")
            break
    }
}
```
