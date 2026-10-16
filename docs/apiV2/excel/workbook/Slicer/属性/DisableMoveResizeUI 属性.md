---
title: "DisableMoveResizeUI 属性"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/Slicer/%E5%B1%9E%E6%80%A7/DisableMoveResizeUI%20%E5%B1%9E%E6%80%A7.html
---
# DisableMoveResizeUI (属性)

返回或设置是否可以使用用户界面来移动指定的切片器或者调整其大小。可读/写。

## 说明

如果无法通过选择用户界面中的边框或手柄来移动切片器或调整其大小，则为 **True**；否则为 **False**。默认值为 **False**。将 **DisableMoveResizeUI** 属性设置为 **True** 只影响用户界面。不禁用通过从代码中设置 **[Slicer](./../../Slicer/Slicer 对象.html)** 对象的属性（如 **Top**、**Left**、**Width** 或 **Height** 属性）来移动切片器或调整其大小。

## 返回值

Boolean

## 示例

```javascript
/*本示例判断活动工作簿上第三个切片器缓存中名为“id”的切片器是否可以使用用户界面来移动该切片器或调整该切片器大小，并通知用户。*/
function test() {
    let slicer = ActiveWorkbook.SlicerCaches(3).Slicers.Item("id")
    if (slicer.DisableMoveResizeUI) {
        console.log("不可以使用用户界面来移动该切片器或调整该切片器大小")
    } else {
        console.log("可以使用用户界面来移动该切片器或调整该切片器大小")
    }
}
```

```javascript
/*本示例将活动工作表上数据透视表的第三个切片器设置为不可以使用用户界面来移动该切片器或调整该切片器大小。*/
function test() {
    let slicer = ActiveSheet.Range("I1").PivotTable.Slicers.Item(3)
    slicer.DisableMoveResizeUI = true
}
```
