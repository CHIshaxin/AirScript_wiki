---
title: "Top 属性"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/Slicer/%E5%B1%9E%E6%80%A7/Top%20%E5%B1%9E%E6%80%A7.html
---
# Top (属性)

返回或设置指定切片器在工作表上相对于单元格 A1 左上角的垂直位置（单位为磅）。可读/写。

## 返回值

Double

## 示例

```javascript
/*本示例设置活动工作簿上第一个切片器缓存中“name”切片器相对于单元格 A1 左上角垂直位置和水平位置的距离。*/
function test() {
    let slicer = ActiveWorkbook.SlicerCaches(1).Slicers.Item("name")
    slicer.Top = 0
    slicer.Left = Range("A1:P1").Width
}
```

```javascript
/*本示例显示活动工作表上第一张数据透视表的“id”切片器相对于 A1 单元格左上角垂直位置的距离。*/
function test() {
    console.log(ActiveSheet.PivotTables(1).Slicers.Item("id").Top)
}
```
