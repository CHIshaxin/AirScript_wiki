---
title: "Range (方法)"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/Shapes/%E6%96%B9%E6%B3%95/Range%20%E6%96%B9%E6%B3%95.html
updated: 1789700916000
hash: B1wGNCK8
---

# Range (方法)

返回一个 **[ShapeRange](./../../ShapeRange/ShapeRange 对象.html)** 对象，它代表 **[Shapes](./../../Shapes/Shapes 对象.html)** 集合中形状的子集。

## 说明

虽然使用 **Range** 属性可返回任意数量的形状，但如果要返回集合中单个成员时，用 **Item** 方法更加简单。例如，`Shapes(1)` 比 `Shapes.Range(1)` 简单。

可为 **Index** 指定整数或字符串数组。例如，以下指令返回用名称指定的两个形状。

```javascript
let arShapes = ["Oval 4", "Rectangle 5"]
let objRange = ActiveSheet.Shapes.Range(arShapes)
```

在 ET 中，不能用此属性返回包含工作表上的所有 **[Shape](./../../Shape/Shape 对象.html)** 对象的 **[ShapeRange](./../../ShapeRange/ShapeRange 对象.html)** 对象。如果要达到该目的，可用下列代码：

```javascript
Worksheets.Item(1).Shapes.SelectAll() //select all shapes
let sr = Selection.ShapeRange //create ShapeRange
```

## 参数

| 属性 | 数据类型 | 必填 | 说明 |
| --- | --- | --- | --- |
| Index | any | 必填 | 包含在该区域中的各单个形状。可以是指定形状索引号的整数、指定形状名称的字符串，也可以是包含整数或字符串的数组。 |

## 示例

```javascript
/*本示例设置第一张工作表中第一个形状和第三个形状的填充图案。*/
function test() {
    let shpRange = Application.Worksheets.Item(1).Shapes.Range([1, 3])
    shpRange.Fill.Patterned(msoPatternHorizontalBrick)
}
```

```javascript
/*本示例设置活动工作表中的第一个形状的填充颜色为红色。*/
function test() {
    let shapes = ActiveSheet.Shapes
    let range = shapes.Range(1)
    range.Fill.ForeColor.RGB = RGB(255, 0, 0)
}
```
