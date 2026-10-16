---
title: "Selected (属性)"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/SlicerItem/%E5%B1%9E%E6%80%A7/Selected%20%E5%B1%9E%E6%80%A7.html
updated: 1789700916000
hash: BqJTWC-N
---

# Selected (属性)

返回或设置是否选择了切片器项。连接到非 OLAP 数据源的切片器，可读/写。连接到 OLAP 数据源的切片器，只读。

## 说明

如果选择了切片器项，则为 **True**；否则为 **False**。

## 返回值

Boolean

## 示例

```javascript
/*本示例选中活动工作簿上第二个切片器缓存中第一个切片器项。*/
function test() {
    ActiveWorkbook.SlicerCaches(2).SlicerItems(1).Selected = true
}
```

```javascript
/*本示例显示活动工作表上数据透视表第一个切片器的缓存中名为“张三”的切片器项是否被选中。*/
function test() {
    let sliItem = ActiveSheet.Range("I1").PivotTable.Slicers.Item(1).SlicerCache.SlicerItems("张三")
    console.log(sliItem.Selected)
}
```
