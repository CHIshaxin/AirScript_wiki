---
title: "SourceName (属性)"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/SlicerItem/%E5%B1%9E%E6%80%A7/SourceName%20%E5%B1%9E%E6%80%A7.html
updated: 1789700916000
hash: CSaTXH7v
---

# SourceName (属性)

返回切片器项的源名称。只读。

## 返回值

Variant

## 示例

```javascript
/*本示例显示第一张工作簿上名为“切片器_score”的切片器缓存中第六个切片器项的源名称。*/
function test() {
    let sliItem = Workbooks.Item(1).SlicerCaches("切片器_score").SlicerItems(6)
    console.log(sliItem.SourceName)
}
```

```javascript
/*本示例显示活动工作表上第一张数据透视表的第一个切片器的缓存中第二个切片器项的源名称。*/
function test() {
    let sliItem = ActiveSheet.PivotTables(1).Slicers.Item(1).SlicerCache.SlicerItems(2)
    console.log(sliItem.SourceName)
}
```
