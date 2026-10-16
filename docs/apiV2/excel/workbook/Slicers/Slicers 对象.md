---
title: "Slicers 对象"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/Slicers/Slicers%20%E5%AF%B9%E8%B1%A1.html
---
# Slicers (对象)

**[Slicer](./../Slicer/Slicer 对象.html)** 对象的集合。

## 说明

每个 **[Slicer](./../Slicer/Slicer 对象.html)** 对象都代表工作簿中的一个切片器。切片器的作用是筛选数据。

使用 **Slicers** 属性可返回 **[Slicers](./../Slicers/Slicers 对象.html)** 集合。

## 示例

```javascript
/*本示例显示活动工作簿中第一个切片器缓存中的切片器数量。*/
function test() {
    console.log(ActiveWorkbook.SlicerCaches(1).Slicers.Count)
}
```

使用 **Slicers** (*index*)（其中 *index* 是切片器的索引号或名称）可返回切片器集合中的单个 **[Slicer](./../Slicer/Slicer 对象.html)** 对象。

```javascript
/*本示例将第一个切片器缓存中第一个切片器的标题更改为“My Slicer”。*/
function test() {
    ActiveWorkbook.SlicerCaches(1).Slicers.Item(1).Caption = "My Slicer"
}
```
