---
title: "Delete 方法"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/Slicer/%E6%96%B9%E6%B3%95/Delete%20%E6%96%B9%E6%B3%95.html
---
# Delete (方法)

删除切片器并将其从关联的 **[Slicers](./../../Slicers/Slicers 对象.html)** 集合中删除。

## 返回值

null

## 示例

```javascript
/*本示例删除活动工作表上第一张数据透视表的第一个切片器。*/
function test() {
    let slicer = ActiveSheet.PivotTables(1).Slicers.Item(1)
    slicer.Delete()
}
```

```javascript
/*本示例判断如果活动工作簿上名为“切片器_date”的切片器缓存的第一个切片器的标题为“date”，则删除该切片器。*/
function test() {
    let slicer = ActiveWorkbook.SlicerCaches("切片器_date").Slicers.Item(1)
    if (slicer.Caption == "date") {
        slicer.Delete()
    }
}
```
