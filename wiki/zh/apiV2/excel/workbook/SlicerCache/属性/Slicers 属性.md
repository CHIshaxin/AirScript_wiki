---
title: "Slicers (属性)"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/SlicerCache/%E5%B1%9E%E6%80%A7/Slicers%20%E5%B1%9E%E6%80%A7.html
updated: 1789700916000
hash: DRwcP9y_
---

# Slicers (属性)

返回一个 **[Slicers](./../../Slicers/Slicers 对象.html)** 集合，其中包含与指定的 **SlicerCache** 关联的 **[Slicer](./../../Slicer/Slicer 对象.html)** 对象的集合。只读

## 返回值

Slicers

## 示例

```javascript
/*本示例显示活动工作簿上第一个切片器缓存中第一个切片器的名称。*/
function test() {
    console.log(ActiveWorkbook.SlicerCaches(1).Slicers(1).Name)
}
```

```javascript
/*本示例显示活动工作簿上名为“切片器_id”的切片器缓存中第一个切片器的高度。*/
function test() {
    console.log(ActiveWorkbook.SlicerCaches("切片器_id").Slicers(1).Height)
}
```
