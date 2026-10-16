---
title: "Group (方法)"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/ShapeRange/%E6%96%B9%E6%B3%95/Group%20%E6%96%B9%E6%B3%95.html
updated: 1789700916000
hash: Q1VBL6uJ
---

# Group (方法)

将指定区域中的形状形成一组。

## 说明

由于一组形状作为单个形状处理，所以创建和分解形状组将更改 **[Shapes](./../../Shapes/Shapes 对象.html)** 集合中的项目数，并更改集合中受影响的项目之后的项目索引号。

**[Range](./../../Range/Range 对象.html)** 对象必须是数据透视表字段的数据区域中的单个单元格。如果试图对多个单元格应用该方法，将会失败（不显示错误消息）。

## 返回值

一个代表组合形状的 **[Shape](./../../Shape/Shape 对象.html)** 对象。

## 示例

```javascript
/*本示例将第一张工作表中新建的ShapeRange对象进行组合*/
function test() {
    let shpRange = Application.Worksheets.Item(1).Shapes.Range([1, 2])
    shpRange.Group()
}
```

```javascript
/*本示例将活动工作表中新建的ShapeRange对象进行组合，并设置组合后的位置*/
function test() {
    let shpRange = ActiveSheet.Shapes.Range([2, 3])
    let group = shpRange.Group()
    group.Left = 100
    group.Top = 100
}
```
