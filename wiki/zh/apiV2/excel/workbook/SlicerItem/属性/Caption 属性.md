---
title: "Caption (属性)"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/SlicerItem/%E5%B1%9E%E6%80%A7/Caption%20%E5%B1%9E%E6%80%A7.html
updated: 1789700916000
hash: D2nBV8-v
---

# Caption (属性)

返回切片器项的标签文本。只读。

## 返回值

String

## 示例

```javascript
/*本示例显示活动工作簿上名为“切片器_date”的切片器缓存中第一项的标签文本。*/
function test() {
    console.log(ActiveWorkbook.SlicerCaches("切片器_date").SlicerItems.Item(1).Caption)
}
```

```javascript
/*本示例将活动工作表上数据透视表第一个切片器的缓存中切片器项的标签文本逐个赋值到单元格中。*/
function test() {
    let sliItems = ActiveSheet.Range("I1").PivotTable.Slicers.Item(1).SlicerCache.SlicerItems
    for (let i = 1; i <= sliItems.Count; i++) {
        Cells.Item(i + 10, 1).Value2 = sliItems.Item(i).Caption
    }
}
```
