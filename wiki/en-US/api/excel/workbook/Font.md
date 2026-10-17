---
title: "Font"
source: https://airsheet.wps.cn/docs/en-US/api/excel/workbook/Font.html
updated: 1789700916000
hash: DNf74izZ
---

# Font

Font properties for cell text: bold, color, size, italic, strikethrough, and underline.

See the lists below for the Font object's properties and methods.

### Properties

| Property | Data type | Description |
| --- | --- | --- |
| [Bold](#bold) | Boolean | Whether the font is bold |
| [Color](#color) | String | Font color |
| [Italic](#italic) | Boolean | Whether the font is italic |
| [Name](#name) | String | Font name |
| [Size](#size) | Number | Font size |
| [Strikethrough](#strikethrough) | Boolean | Whether strikethrough is applied |
| [Underline](#underline) | [Enum.XlUnderlineStyle](Enum.md#xlunderlinestyle) | Underline style |

## Bold

Gets or sets bold formatting.

### Data type

Boolean - Whether the font is bold

### Example

```js
// Font object
let font = Application.Range('A1').Font

// Print whether font is bold
console.log('Bold:', font.Bold)

// Set bold
font.Bold = true
```

## Color

Gets or sets the font color.

### Data type

String - Font color as hexadecimal RGB

### Example

```js
// Font object
let font = Application.Range('A1').Font

// Print font color
console.log('Color:', font.Color)

// Set font color
font.Color = '#eb5451'
```

## Italic

Gets or sets italic formatting.

### Data type

Boolean - Whether the font is italic

### Example

```js
// Font object
let font = Application.Range('A1').Font

// Print whether font is italic
console.log('Italic:', font.Italic) // Italic: false

// Set italic
font.Italic = true
```

## Name

Gets or sets the font name.

### Data type

String - Font name

### Example

```js
// Set A1 font to Microsoft YaHei
Range('A1').Font.Name = 'Microsoft YaHei'
```

## Size

Gets or sets font size.

### Data type

Number - Font size

### Example

```js
// Font object
let font = Application.Range('A1').Font

// Print font size
console.log(font.Size) //12

// Set font size
font.Size = 30
```

## Strikethrough

Gets or sets strikethrough.

### Data type

Boolean - Whether strikethrough is applied

### Example

```js
// Font object
let font = Application.Range('A1').Font

// Print whether strikethrough is on
console.log('Strikethrough:', font.Strikethrough) // Strikethrough: false

// Enable strikethrough
font.Strikethrough = true
```

## Underline

Gets or sets underline style.

### Data type

[Enum.XlUnderlineStyle](Enum.md#xlunderlinestyle) - Underline type

For example, single underline: `Application.Enum.XlUnderlineStyle.xlUnderlineStyleSingle`

### Example

```js
// Font object
const font = Application.Range('A1').Font

// Print underline setting
console.log('Underline:', font.Underline) // Underline: -4142

const underlineStatus = Application.Enum.XlUnderlineStyle.xlUnderlineStyleSingle

// Single underline
font.Underline = underlineStatus
```
