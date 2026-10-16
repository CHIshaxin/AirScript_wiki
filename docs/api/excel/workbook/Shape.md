---
title: "图形(Shape)"
source: https://airsheet.wps.cn/docs/api/excel/workbook/Shape.html
---
# Shape

某个工作表上的单个 Shape 图形对象

Shape 对象的具体属性和方法请参阅下方的列表。

### 属性列表

| 属性名 | 数据类型 | 简介 |
| --- | --- | --- |
| ID | String | 图形 ID |
| Name | String | 图形名称 |
| Title | String | 图形标题 |
| Chart | [Chart](Chart.md) | 图表对象 |

### 方法列表

| 方法名 | 返回类型 | 简介 |
| --- | --- | --- |
| Delete | undefined | 删除Shape |

## ID

单个图形 ID

### 数据类型

String - 图形对象的 ID

### 示例

```js
// 图形对象集合
const shape = Application.ActiveSheet.Shapes.Item(1)

// 打印该图形对象的ID
console.log(shape.ID)  //94646582190480
```

## Name

单个图形名称

### 数据类型

String - 图形对象的名称

### 示例

```js
// 图形对象集合
const shape = Application.ActiveSheet.Shapes.Item(1)

// 打印图形对象的名称
console.log(shape.Name) //图表 1
```

## Title

单个图形标题

该对象只能设置值，无法读取值

### 数据类型

String - 图形对象的标题

### 示例

```js
const shape = Application.ActiveSheet.Shapes.Item(1)

// 设置图形对象的标题,该属性只写,效果需要自行观察图表对象
shape.Title = 'WebOffice'
```

## Chart

单个图表对象

### 数据类型

[Chart](Chart.md) - 单个图表对象

### 示例

```js
const shape = Application.ActiveSheet.Shapes.Item(1)
shape.Chart.HasTitle = false // 不展示标题
```

## Delete()

删除图标

### 数据类型

undefined

### 示例

```js
const shape = Application.ActiveSheet.Shapes.Item(1)
shape.Delete()
```
