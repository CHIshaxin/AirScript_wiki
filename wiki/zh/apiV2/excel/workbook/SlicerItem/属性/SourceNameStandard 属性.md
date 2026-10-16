---
title: "SourceNameStandard (属性)"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/SlicerItem/%E5%B1%9E%E6%80%A7/SourceNameStandard%20%E5%B1%9E%E6%80%A7.html
updated: 1789700916000
hash: Cl7jUrDy
---

# SourceNameStandard (属性)

返回以标准英语（美国）格式设置表示的切片器项的源名称。只读。

## 说明

当切片器项有本地化版本且其 **SourceNameStandard** 属性值不同于 **SourceName** 属性值时，将使用此属性，例如对于日期格式。

## 返回值

String

## 示例

```javascript
/*本示例显示以标准英语（美国）格式表示的第一张工作簿上名为“切片器_date”的切片器缓存中第一个切片器项的源名称。*/
function test() {
    let sliItem = Workbooks.Item(1).SlicerCaches("切片器_date").SlicerItems(1)
    console.log(sliItem.SourceNameStandard)
}
```

```javascript
/*本示例显示以标准英语（美国）格式表示的活动工作表上数据透视表的第一个切片器的缓存的第一个切片器项的源名称。*/
function test() {
    let sliItem = ActiveSheet.Range("I1").PivotTable.Slicers.Item(1).SlicerCache.SlicerItems(1)
    console.log(sliItem.SourceNameStandard)
}
```
