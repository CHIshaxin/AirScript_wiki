---
title: "Height (属性)"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/Slicer/%E5%B1%9E%E6%80%A7/Height%20%E5%B1%9E%E6%80%A7.html
updated: 1789700916000
hash: C_-6AUcl
---

# Height (属性)

返回或设置指定切片器的高度（单位磅）。可读/写。

## 返回值

Double

## 示例

```javascript
/*本示例显示活动工作簿上第四个切片器缓存中“age”切片器的高度。*/
function test() {
    console.log(ActiveWorkbook.SlicerCaches(4).Slicers.Item("age").Height)
}
```

```javascript
/*本示例将第一张工作表上第一张数据透视表的第一个切片器的高度设置为 200 磅。*/
function test() {
    let slicer = Worksheets.Item(1).PivotTables(1).Slicers.Item(1)
    slicer.Height = 200
}
```
