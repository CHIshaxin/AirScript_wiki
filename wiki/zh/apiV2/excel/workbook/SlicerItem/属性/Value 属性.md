---
title: "Value (属性)"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/SlicerItem/%E5%B1%9E%E6%80%A7/Value%20%E5%B1%9E%E6%80%A7.html
updated: 1789700916000
hash: BePWa0vh
---

# Value (属性)

返回切片器项的值。只读。

## 返回值

String

## 示例

```javascript
/*本示例显示活动工作簿上名为“切片器_name”的切片器缓存中第三个切片器项的值。*/
function test() {
    console.log(ActiveWorkbook.SlicerCaches("切片器_name").SlicerItems(3).Value)
}
```

```javascript
/*本示例显示活动工作表上数据透视表的第一个切片器的缓存中第五个切片器项的值。*/
function test() {
    let sliItem = ActiveSheet.Range("I1").PivotTable.Slicers.Item(1).SlicerCache.SlicerItems(5)
    console.log(`第五个切片器项的值：${sliItem.Value}`)
}
```
