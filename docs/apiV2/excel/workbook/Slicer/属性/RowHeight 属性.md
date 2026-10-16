---
title: "RowHeight 属性"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/Slicer/%E5%B1%9E%E6%80%A7/RowHeight%20%E5%B1%9E%E6%80%A7.html
---
# RowHeight (属性)

返回或设置指定切片器中每一行的高度（单位为磅）。可读/写。

## 说明

此属性不影响切片器控件的高度。使用 **[Slicer](./../../Slicer/Slicer 对象.html)** 对象的 **Height** 属性可返回或设置切片器的高度。

## 返回值

Double

## 示例

```javascript
/*本示例显示活动工作表上第一张数据透视表的第二个切片器每一行的高度。*/
function test() {
    console.log(ActiveSheet.PivotTables(1).Slicers.Item(2).RowHeight)
}
```

```javascript
/*本示例将活动工作簿上名为“切片器_name”的切片器缓存中第一个切片器每一行的高度设置为 25 磅。*/
function test() {
    let slicer = ActiveWorkbook.SlicerCaches("切片器_name").Slicers.Item(1)
    slicer.RowHeight = 25
}
```
