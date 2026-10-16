---
title: "Width 属性"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/Slicer/%E5%B1%9E%E6%80%A7/Width%20%E5%B1%9E%E6%80%A7.html
---
# Width (属性)

获取或设置指定切片器的宽度（以磅为单位）。可读/写。

## 返回值

Double

## 示例

```javascript
/*本示例设置活动工作簿上第三个切片器缓存中“id”切片器的宽度。*/
function test() {
    let slicer = ActiveWorkbook.SlicerCaches(3).Slicers.Item("id")
    slicer.Width = (Range("F1").Width) * 3
}
```

```javascript
/*本示例显示活动工作表上第一张数据透视表的第二个切片器的宽度。*/
function test() {
    console.log(ActiveSheet.PivotTables(1).Slicers.Item(2).Width)
}
```
