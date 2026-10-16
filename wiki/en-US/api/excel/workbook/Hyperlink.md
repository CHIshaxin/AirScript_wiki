---
title: "Hyperlink"
source: https://airsheet.wps.cn/docs/en-US/api/excel/workbook/Hyperlink.html
updated: 1789700916000
hash: -jjhco9f
---

# Hyperlink

A single hyperlink object. It has a jump address and display text; the two do not have to be the same.

See the lists below for the Hyperlink object's properties and methods.

### Properties

| Property | Data type | Description |
|---|---|---|
| [Address](#address) | String | Hyperlink target URL |
| [TextToDisplay](#texttodisplay) | String | Hyperlink display text |

## Address

Gets or sets the hyperlink address.

### Data type

String - Hyperlink target URL.

### Example

```js
// Get hyperlink object
const hyperlink = Application.ActiveSheet.Hyperlinks.Item(1)
// Print hyperlink address
console.log(hyperlink.Address) // https://www.wps.com
hyperlink.Address = "https://www.wps.com"
```

## TextToDisplay

Gets or sets the hyperlink display text.

### Data type

String - Hyperlink display text.

### Example

```js
// Get hyperlink object
const hyperlink = Application.ActiveSheet.Hyperlinks.Item(1)
// Print display text
console.log(hyperlink.TextToDisplay) // WPS Docs website
hyperlink.TextToDisplay = "WPS Docs"
```
