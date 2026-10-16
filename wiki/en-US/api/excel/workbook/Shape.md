---
title: "Shape"
source: https://airsheet.wps.cn/docs/en-US/api/excel/workbook/Shape.html
updated: 1789700916000
hash: UFrE3Zbo
---

# Shape

A single Shape graphic object on a worksheet.

See the lists below for the Shape object's properties and methods.

### Properties

| Property | Data type | Description |
|---|---|---|
| [ID](#id) | String | Shape ID |
| [Name](#name) | String | Shape name |
| [Title](#title) | String | Shape title |
| [Chart](#chart) | [Chart](/docs/en-US/api/excel/workbook/Chart.html) | Chart object |

### Methods

| Method | Return type | Description |
|---|---|---|
| [Delete](#delete) | undefined | Deletes the shape |

## ID

Shape id.

### Data type

String - Shape object id

### Example

```js
// Shape collection
const shape = Application.ActiveSheet.Shapes.Item(1)

// Print shape id
console.log(shape.ID)  //94646582190480
```

## Name

Shape name.

### Data type

String - Shape object name

### Example

```js
// Shape collection
const shape = Application.ActiveSheet.Shapes.Item(1)

// Print shape name
console.log(shape.Name) // Chart 1
```

## Title

Shape title.

Write-only; cannot be read.

### Data type

String - Shape title

### Example

```js
const shape = Application.ActiveSheet.Shapes.Item(1)

// Set title (write-only; observe effect on chart)
shape.Title = 'WebOffice'
```

## Chart

Single chart object.

### Data type

[Chart](/docs/en-US/api/excel/workbook/Chart.html) - Chart object

### Example

```js
const shape = Application.ActiveSheet.Shapes.Item(1)
shape.Chart.HasTitle = false // hide title
```

## Delete()

Deletes the shape.

### Data type

undefined

### Example

```js
const shape = Application.ActiveSheet.Shapes.Item(1)
shape.Delete()
```
