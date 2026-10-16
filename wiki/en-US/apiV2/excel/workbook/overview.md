---
title: "Workbook API"
source: https://airsheet.wps.cn/docs/en-US/apiV2/excel/workbook/overview.html
updated: 1789700916000
hash: CgK6ZqLV
---

# Workbook API

The Workbook API provides complete control over Excel worksheets, including cell operations, formatting, chart creation, data filtering, and more.

## Core Objects

- **Application** - The application object providing global methods and properties
- **Workbook** - The workbook object managing the worksheet collection
- **Worksheet** - The worksheet object for operating on individual sheets
- **Range** - The range object for manipulating cell ranges
- **Chart** - The chart object for creating and manipulating charts

## Quick Start

```javascript
// Get the current active sheet
const sheet = ActiveSheet;

// Read cell value
const value = sheet.Cells(1, 1).Value2;

// Set cell value
sheet.Cells(1, 1).Value2 = "Hello World";

// Get a range
const range = sheet.Range("A1:B10");
```

Please select a specific API object from the sidebar navigation for detailed documentation.
