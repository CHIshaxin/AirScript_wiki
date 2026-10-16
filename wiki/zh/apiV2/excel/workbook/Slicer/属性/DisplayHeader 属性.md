---
title: "DisplayHeader (属性)"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/Slicer/%E5%B1%9E%E6%80%A7/DisplayHeader%20%E5%B1%9E%E6%80%A7.html
updated: 1789700916000
hash: B57xknV5
---

# DisplayHeader (属性)

返回或设置显示切片器 **Caption** 的页眉是否可见。可读/写。

## 说明

如果页眉可见，则为 **True**；否则为 **False**。默认值为 **True**。

## 返回值

Boolean

## 示例

```javascript
/*本示例显示活动工作表上数据透视表的第三个切片器的标题页眉是否可见。*/
function test() {
    console.log(ActiveSheet.Range("I1").PivotTable.Slicers.Item(3).DisplayHeader)
}
```

```javascript
/*本示例将活动工作簿中第三个切片器缓存中名为“id”的切片器的标题页眉设置为不可见。*/
function test() {
    let slicer = ActiveWorkbook.SlicerCaches(3).Slicers.Item("id")
    slicer.DisplayHeader = false
}
```
