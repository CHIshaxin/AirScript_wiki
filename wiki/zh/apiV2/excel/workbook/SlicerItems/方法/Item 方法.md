---
title: "Item (方法)"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/SlicerItems/%E6%96%B9%E6%B3%95/Item%20%E6%96%B9%E6%B3%95.html
updated: 1789700916000
hash: Cr4HzkeL
---

# Item (方法)

从集合中返回一个 **[SlicerItem](./../../SlicerItem/SlicerItem 对象.html)** 对象。只读。

## 参数

| 属性 | 数据类型 | 必填 | 说明 |
| --- | --- | --- | --- |
| Index | any | 必填 | 对象的名称或索引号。 |

## 返回值

SlicerItem

## 示例

```javascript
/*本示例显示活动工作簿上第二个切片器缓存中第三个切片器项的源名称。*/
function test() {
    console.log(ActiveWorkbook.SlicerCaches(2).SlicerItems.Item(3).SourceName)
}
```

```javascript
/*本示例选中活动工作表上数据透视表第一个切片器的缓存中名为“张三”的切片器项。*/
function test() {
    let sliItem = ActiveSheet.Range("I1").PivotTable.Slicers.Item(1).SlicerCache.SlicerItems.Item("张三")
    sliItem.Selected = true
}
```
