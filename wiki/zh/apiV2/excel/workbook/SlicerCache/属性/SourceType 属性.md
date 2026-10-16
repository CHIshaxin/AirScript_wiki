---
title: "SourceType (属性)"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/SlicerCache/%E5%B1%9E%E6%80%A7/SourceType%20%E5%B1%9E%E6%80%A7.html
updated: 1789700916000
hash: XzkDekkg
---

# SourceType (属性)

返回切片器所连接到的数据源的类型。只读

## 说明

只支持 **xlDatabase** 和 **xlExternal** 数据源类型。

## 返回值

XlPivotTableSourceType

## 示例

```javascript
/*本示例显示活动工作簿上名为“切片器_id”的切片器缓存所连接到的数据源的类型是否为 xlDatabase。*/
function test() {
    console.log(ActiveWorkbook.SlicerCaches("切片器_id").SourceType == xlDatabase)
}
```

```javascript
/*本示例判断第一张工作簿上第二个切片器缓存所连接到的数据源的类型是否为xlPivotTable，并通知用户。*/
function test() {
    let slcCache = Workbooks.Item(1).SlicerCaches(2)
    if (slcCache.SourceType == xlPivotTable) {
        console.log("数据源的类型是 xlPivotTable")
    } else {
        console.log("数据源的类型不是 xlPivotTable")
    }
}
```
