---
title: "ColumnWidth (属性)"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/Slicer/%E5%B1%9E%E6%80%A7/ColumnWidth%20%E5%B1%9E%E6%80%A7.html
updated: 1789700916000
hash: Bg4a-6q_
---

# ColumnWidth (属性)

获取或设置切片器控件中每一列的宽度（以磅为单位）。可读/写。

## 说明

**[Slicer](./../../Slicer/Slicer 对象.html)** 对象的 **Width** 属性根据 **ColumnWidth** 属性和 **NumberOfColumns** 属性的值自动调整。

## 返回值

Double

## 示例

```javascript
/*本示例显示活动工作簿上第二个切片器缓存中第一个切片器每一列的宽度。*/
function test() {
    console.log(ActiveWorkbook.SlicerCaches(2).Slicers.Item(1).ColumnWidth)
}
```

```javascript
/*本示例设置工作表 Sheet1 上第一张数据透视表的第一个切片器每一列的宽度，并显示设置前后的该切片器每一列宽度。*/
function test() {
    let slicer = Worksheets.Item("Sheet1").PivotTables(1).Slicers.Item(1)
    console.log(`设置前切片器每一列的宽度：${slicer.ColumnWidth} 磅`)
    slicer.ColumnWidth = 120
    console.log(`设置后切片器每一列的宽度：${slicer.ColumnWidth} 磅`)
}
```
