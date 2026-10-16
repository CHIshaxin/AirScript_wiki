---
title: "Name (属性)"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/SlicerItem/%E5%B1%9E%E6%80%A7/Name%20%E5%B1%9E%E6%80%A7.html
updated: 1789700916000
hash: DYPalZAE
---

# Name (属性)

返回切片器项的名称。只读。

## 返回值

String

## 示例

```javascript
/*本示例显示活动工作簿上名为“切片器_date”的切片器缓存中第三个切片器项的名称。*/
function test() {
    console.log(ActiveWorkbook.SlicerCaches("切片器_date").SlicerItems.Item(3).Name)
}
```

```javascript
/*本示例显示活动工作表上第一张数据透视表的第一个切片器的缓存中第四个切片器项的名称。*/
function test() {
    let sliItem = ActiveSheet.PivotTables(1).Slicers.Item(1).SlicerCache.SlicerItems.Item(4)
    console.log(`第四个切片器项的名称：${sliItem.Name}`)
}
```
