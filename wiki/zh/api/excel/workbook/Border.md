---
title: "Border"
source: https://airsheet.wps.cn/docs/api/excel/workbook/Border.html
updated: 1789700916000
hash: C22ddnFB
---

# Border

边框对象，Borders 集合里的某一边框

Border 对象的具体属性和方法请参阅下方的列表。

### 属性列表

| 属性名 | 数据类型 | 简介 |
|---|---|---|
| [Color](#color) | Number | 边框的颜色 |
| [Weight](#weight) | [Enum.XlBorderWeight](/docs/api/excel/workbook/Enum.html#xlborderweight) | 边框的粗细 |
| [LineStyle](#linestyle) | [Enum.XlLineStyle](/docs/api/excel/workbook/Enum.html#xllinestyle) | 边框的线条样式 |

## Color

边框颜色

注意：获取边框颜色时，需要指定具体的边框，即枚举值 [Enum.XlBordersIndex](/docs/api/excel/workbook/Enum.html#xlbordersindex) 不能是 xlAll、xlOutside、xlInside 等

### 数据类型

Number - 边框颜色

### 示例

```js
// 区域对象
const borders = Application.Range('A1').Borders

// 底部边框对象
let border = borders.Item(Application.Enum.XlBordersIndex.xlEdgeBottom)

// 打印边框颜色
console.log(border.Color) // {"private":{"isAuto":false,"rgbValue":"#000000","tint":0,"type":"ectAUTO"},"rgbValue":"#000000"}

// 将整个边框颜色设置成绿色
border = borders.Item(Application.Enum.XlBordersIndex.xlOutside)
border.Color = '#00FF00'
```

## Weight

边框的粗细

### 数据类型

[Enum.XlBorderWeight](/docs/api/excel/workbook/Enum.html#xlborderweight)- 边框的粗细

该对象只能设置值，无法读取值

### 示例

```js
// 区域对象
let borders = Application.Range('A3').Borders

// 单个边框对象
let border = borders.Item(Application.Enum.XlBordersIndex.xlOutside)

// 设置边框的粗细
border.Weight = Application.Enum.XlBorderWeight.xlThick
```

## LineStyle

边框的线条样式

### 数据类型

[Enum.XlLineStyle](/docs/api/excel/workbook/Enum.html#xllinestyle)- 边框的线条样式

该对象只能设置值，无法读取值

### 示例

```js
// 区域对象
let borders = Application.Range('A1').Borders

// 底部边框对象
let border = borders.Item(Application.Enum.XlBordersIndex.xlEdgeBottom)

// 将整个边框样式设置成双线
border = borders.Item(Application.Enum.XlBordersIndex.xlOutside)
border.LineStyle = Application.Enum.XlLineStyle.xlDouble
```
