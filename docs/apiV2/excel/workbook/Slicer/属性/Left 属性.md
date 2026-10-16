---
title: "Left 属性"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/Slicer/%E5%B1%9E%E6%80%A7/Left%20%E5%B1%9E%E6%80%A7.html
---
# Left (属性)

返回或设置指定切片器在工作表上相对于单元格 A1 左上角的水平位置（单位为磅）。可读/写。

## 返回值

Double

## 示例

```javascript
/*本示例显示活动工作表上第一张数据透视表的第三个切片器到单元格 A1 左上角水平位置的距离。*/
function test() {
    console.log(ActiveSheet.PivotTables(1).Slicers.Item(3).Left)
}
```

```javascript
/*本示例将活动工作簿上第一个切片器缓存中“name”切片器的高度设置为 160 磅，并将该切片器到单元格 A1 左上角水平位置的距离设置为 10 磅。*/
function test() {
    let slicer = ActiveWorkbook.SlicerCaches(1).Slicers.Item("name")
    slicer.Height = 160
    slicer.Left = 10
}
```
