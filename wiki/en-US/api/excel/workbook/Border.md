---
title: "Border"
source: https://airsheet.wps.cn/docs/en-US/api/excel/workbook/Border.html
updated: 1789700916000
hash: nYYArOtl
---

# Border

Border object: one border in a Borders collection.

See the lists below for the Border object's properties and methods.

### Properties

| Property | Data type | Description |
|---|---|---|
| [Color](#color) | Number | Border color |
| [Weight](#weight) | [Enum.XlBorderWeight](/docs/en-US/api/excel/workbook/Enum.html#xlborderweight) | Border weight |
| [LineStyle](#linestyle) | [Enum.XlLineStyle](/docs/en-US/api/excel/workbook/Enum.html#xllinestyle) | Border line style |

## Color

Border color.

Note: When reading border color, specify a concrete border; [Enum.XlBordersIndex](/docs/en-US/api/excel/workbook/Enum.html#xlbordersindex) must not be xlAll, xlOutside, xlInside, etc.

### Data type

Number - Border color

### Example

```js
// Range borders
const borders = Application.Range('A1').Borders

// Bottom border
let border = borders.Item(Application.Enum.XlBordersIndex.xlEdgeBottom)

// Print border color
console.log(border.Color) // {"private":{"isAuto":false,"rgbValue":"#000000","tint":0,"type":"ectAUTO"},"rgbValue":"#000000"}

// Set full outside border to green
border = borders.Item(Application.Enum.XlBordersIndex.xlOutside)
border.Color = '#00FF00'
```

## Weight

Border weight.

### Data type

[Enum.XlBorderWeight](/docs/en-US/api/excel/workbook/Enum.html#xlborderweight) - Border weight

Write-only; cannot be read.

### Example

```js
// Range borders
let borders = Application.Range('A3').Borders

// Single border
let border = borders.Item(Application.Enum.XlBordersIndex.xlOutside)

// Set weight
border.Weight = Application.Enum.XlBorderWeight.xlThick
```

## LineStyle

Border line style.

### Data type

[Enum.XlLineStyle](/docs/en-US/api/excel/workbook/Enum.html#xllinestyle) - Line style

Write-only; cannot be read.

### Example

```js
// Range borders
let borders = Application.Range('A1').Borders

// Bottom border
let border = borders.Item(Application.Enum.XlBordersIndex.xlEdgeBottom)

// Set outside border to double line
border = borders.Item(Application.Enum.XlBordersIndex.xlOutside)
border.LineStyle = Application.Enum.XlLineStyle.xlDouble
```
