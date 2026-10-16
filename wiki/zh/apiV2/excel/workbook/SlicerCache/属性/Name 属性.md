---
title: "Name (属性)"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/SlicerCache/%E5%B1%9E%E6%80%A7/Name%20%E5%B1%9E%E6%80%A7.html
updated: 1789700916000
hash: 8Gl4ZkU_
---

# Name (属性)

返回或设置切片器缓存的名称。

## 说明

切片器缓存的名称在工作簿命名空间中必须是唯一的。默认情况下，分配给切片器缓存的名称是 `Slicer_` 后跟与该切片器缓存关联的数据透视表字段的名称。例如，如果切片器与数据透视表中的“Product Category”字段关联，则默认名称为 `Slicer_Product_Category`（字段名中的任何空格替换为下划线字符）。如果同一工作簿中有多个“Product Category”字段与某个切片器关联，或者工作簿中有某个其他命名实体也名为 `Slicer_Product_Category`，那么 ET 将在该名称后面附加一个编号来产生唯一名称，如 `Slicer_Product_Category1`。

## 返回值

String

## 示例

```javascript
/*本示例显示活动工作簿第一个切片器缓存的名称。*/
function test() {
    console.log(ActiveWorkbook.SlicerCaches(1).Name)
}
```

```javascript
/*本示例设置活动工作簿第三个切片器缓存的名称并显示。*/
function test() {
    ActiveWorkbook.SlicerCaches(3).Name = "切片器_nameid"
    console.log(ActiveWorkbook.SlicerCaches(3).Name)
}
```
