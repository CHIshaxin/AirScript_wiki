---
title: "Slicer (对象)"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/Slicer/Slicer%20%E5%AF%B9%E8%B1%A1.html
updated: 1789700916000
hash: CYsk-CyU
---

# Slicer (对象)

代表工作簿中的一个切片器。

## 说明

每个 **[Slicer](./../Slicer/Slicer 对象.html)** 对象都代表工作簿中的一个切片器。切片器用于筛选数据透视表或 OLAP 数据源中的数据。

使用 **Add** 方法可将 **[Slicer](./../Slicer/Slicer 对象.html)** 对象添加到 **[Slicers](./../Slicers/Slicers 对象.html)** 集合。若要访问表示切片器中当前选中的按钮的 **[SlicerItem](./../SlicerItem/SlicerItem 对象.html)** 对象，请使用 **[Slicer](./../Slicer/Slicer 对象.html)** 对象的 **ActiveItem** 属性。

## 示例

```javascript
/*本示例将活动工作簿上第一个切片器缓存中第一个切片器的标题更改为“My Slicer”。*/
function test() {
    ActiveWorkbook.SlicerCaches(1).Slicers.Item(1).Caption = "My Slicer"
}
```

```javascript
/*本示例将活动工作簿上第一个切片器缓存中第一个切片器的宽度设置为 200 磅。*/
function test() {
    ActiveWorkbook.SlicerCaches(1).Slicers.Item(1).Width = 200
}
```
