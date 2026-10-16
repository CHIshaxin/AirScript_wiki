---
title: "Shape (属性)"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/Slicer/%E5%B1%9E%E6%80%A7/Shape%20%E5%B1%9E%E6%80%A7.html
updated: 1789700916000
hash: B8FA18ef
---

# Shape (属性)

返回与指定的切片器关联的 **[Shape](./../../Shape/Shape 对象.html)** 对象。只读。

## 说明

切片器的某些属性（如 **AlternativeText** 属性）只能从与该切片器关联的 **[Shape](./../../Shape/Shape 对象.html)** 对象访问。

## 返回值

Shape

## 示例

```javascript
/*本示例显示活动工作表上第一张数据透视表的第四个切片器的 Shape 属性关联的可选文字的说明。*/
function test() {
    let slicer = ActiveSheet.PivotTables(1).Slicers.Item(4)
    console.log(slicer.Shape.AlternativeText)
}
```

```javascript
/*本示例设置活动工作簿上名为“切片器_date”的切片器缓存中第一个切片器的 Shape 属性的标题。*/
function test() {
    let slicer = ActiveWorkbook.SlicerCaches("切片器_date").Slicers.Item(1)
    slicer.Shape.Title = "slicer_date"
}
```
