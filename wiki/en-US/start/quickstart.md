---
title: "Getting Started"
source: https://airsheet.wps.cn/docs/en-US/start/quickstart.html
updated: 1789700916000
hash: faQjqjOR
---

# Getting Started

## Open AirScript Editor

1. Create a new spreadsheet on the WPS Docs homepage and open it to experience AirScript.
2. After opening the spreadsheet, click **Efficiency** - **AirScript Editor Tool** in the top menu to open the editor page.
3. Run the examples below one by one and observe the results to quickly get started with AirScript.

## HelloWorld

### console.log()

```js
console.log("hello world!")
```

### console.error()

```js
console.error("hello world!")
```

## Document Operations

### Iterate through SheetNames, switch to corresponding Sheet, read cells, modify cells

```js
// Iterate and print all worksheet names
let sheets = Application.Sheets
for (let i = 0; i < sheets.Count; i++) {
    let sheet = sheets.Item(i + 1)
    console.log(sheet.Name) // Print each worksheet name
}

// Print the name of the currently active Sheet
console.log(Application.ActiveSheet.Name) // Print the active Sheet name

// Get the A1 cell of the currently active worksheet
let A1 = Application.ActiveSheet.Range("A1")

// Print A1 cell content
console.log(A1.Text) // Print the content of A1 cell

// Modify A1 cell content
A1.Value2 = "bar"
console.log(A1.Text) // Print the modified A1 cell content

// Change A1 cell background color to yellow (255,255,0)
A1.Interior.Color = RGB(255,255,0) // Highlight A1 cell
```

### Add a specified suffix to all non-empty cells in the document

```js
// Get the used range through Application.ActiveSheet.UsedRange
const usedRange = Application.ActiveSheet.UsedRange
const startRow = usedRange.Row // Get starting row
const startCol = usedRange.Column // Get starting column
const endRow = startRow + usedRange.Rows.Count // Get ending row
const endCol = startCol + usedRange.Columns.Count // Get ending column

for (let i = startRow; i < endRow; i++) {
    const row = Application.ActiveSheet.Rows(i) // Determine row
    for (let j = startCol; j < endCol; j++) {
        const rg = row.Columns(j) // Specify column from row object to determine cell
        const text = rg.Text
        // If cell is not empty, add foo suffix
        if (text !== '') {
            rg.Value2 = text + 'foo' // Modify cell content
        }
    }
}

// Iterate through each cell in UsedRange and print its content
for (let i = startRow; i < endRow; i++) {
    for (let j = startCol; j < endCol; j++) {
        const cell = Application.ActiveSheet.Cells(i, j)
        console.log(cell.Text) // Print all cell data
    }
}
```

### Create New Sheet

```js
// Current worksheet name
const currentSheetName = Application.ActiveSheet.Name

// Add two worksheets with default names after the current Sheet
Application.Sheets.Add(null, Application.Sheets(currentSheetName), 2)

// Add a Sheet named 'New Worksheet (Right)' after the current Sheet
const newSheetRight = Application.Sheets.Add(null, Application.Sheets(currentSheetName), 1)
newSheetRight.Name = 'New Worksheet (Right)'

// Add a Sheet named 'New Worksheet (Left)' before the current Sheet
const newSheetLeft = Application.Sheets.Add(Application.Sheets(currentSheetName), null, 1)
newSheetLeft.Name = 'New Worksheet (Left)'
```

## More

For more real-world examples, try [Example Cases](../example/excel.md).
