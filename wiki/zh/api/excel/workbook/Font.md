---
title: "Font"
source: https://airsheet.wps.cn/docs/api/excel/workbook/Font.html
updated: 1789700916000
hash: Dsu_CheN
---

# Font

单元格内字体的属性，包括加粗，颜色，大小，斜体，删除线和下划线。

Font 对象的具体属性和方法请参阅下方的列表。

### 属性列表

| 属性名 | 数据类型 | 简介 |
| --- | --- | --- |
| [Bold](#bold) | Boolean | 字体是否加粗 |
| [Color](#color) | String | 字体的颜色 |
| [Italic](#italic) | Boolean | 字体是否斜体 |
| [Name](#name) | String | 字体的名称 |
| [Size](#size) | Number | 字体的大小 |
| [Strikethrough](#strikethrough) | Boolean | 字体是否有删除线 |
| [Underline](#underline) | [Enum.XlUnderlineStyle](Enum.md#xlunderlinestyle) | 字体是否有下划线 |

## Bold

获取或者设置是否加粗

### 数据类型

Boolean - 字体是否加粗

### 示例

```js
// 字体对象
let font = Application.Range('A1').Font

// 打印字体是否加粗
console.log('字体是否加粗：', font.Bold)

// 设置字体加粗
font.Bold = true
```

## Color

获取或者设置字体的颜色

### 数据类型

String - 字体的颜色，16 进制的 RGB 格式

### 示例

```js
// 字体对象
let font = Application.Range('A1').Font

// 打印字体颜色
console.log('字体颜色：', font.Color)

// 设置字体颜色
font.Color = '#eb5451'
```

## Italic

获取或者设置字体是否是斜体

### 数据类型

Boolean - 字体是否是斜体

### 示例

```js
// 字体对象
let font = Application.Range('A1').Font

// 打印字体是否斜体
console.log('字体是否为斜体：', font.Italic) // 字体是否为斜体：false

// 设置字体斜体
font.Italic = true
```

## Name

获取或者设置字体的名称

### 数据类型

String - 字体名

### 示例

```js
// 设置A1单元格的字体为微软雅黑
Range('A1').Font.Name = '微软雅黑'
```

## Size

设置和获取字体的大小

### 数据类型

Number - 字体大小

### 示例

```js
// 字体对象
let font = Application.Range('A1').Font

// 打印字体大小
console.log(font.Size) //12

// 设置字体大小
font.Size = 30
```

## Strikethrough

获取或者设置字体是否有删除线

### 数据类型

Boolean - 字体是否有删除线

### 示例

```js
// 字体对象
let font = Application.Range('A1').Font

// 打印字体是否有删除线
console.log('字体是否有删除线：', font.Strikethrough) // 字体是否有删除线：false

// 设置字体删除线
font.Strikethrough = true
```

## Underline

获取或者设置字体下划线

### 数据类型

[Enum.XlUnderlineStyle](Enum.md#xlunderlinestyle) - 字体下划线的类型

例如：设置单下划线即设置为 `Application.Enum.XlUnderlineStyle.xlUnderlineStyleSingle`

### 示例

```js
// 字体对象
const font = Application.Range('A1').Font

// 打印字体是否设置下划线
console.log('字体是否有下划线：', font.Underline) //字体是否有下划线：-4142

const underlineStatus = Application.Enum.XlUnderlineStyle.xlUnderlineStyleSingle

// 设置字体有单下划线
font.Underline = underlineStatus
```
