---
title: "Spreadsheet Examples"
source: https://airsheet.wps.cn/docs/en-US/example/excel.html
updated: 1789700916000
hash: DGqD6TFU
---

# Spreadsheet Examples

To help you understand AirScript and quickly get started with script development, we have displayed the source code of some features from **Start** - **Quick Tools** in the spreadsheet. We hope that after reading these real-world example codes, you can effectively use the script editor to solve problems.

## Clear Formulas and Keep Values Only

Manually select one or more cells, then execute the script. After execution, it will batch clear the formulas and keep only the values.

```js
const API = Application

// User selection
const selection = API.Selection
// Get the used range of the user's active workbook
const usedRange = API.ActiveSheet.UsedRange

// Get the minimum intersection of both, i.e., the minimum set to traverse
const rowFrom = Math.max(selection.Row, usedRange.Row)
const rowTo = Math.min(selection.Row + selection.Rows.Count - 1, usedRange.Row + usedRange.Rows.Count - 1)
const colFrom = Math.max(selection.Column, usedRange.Column)
const colTo = Math.min(selection.Column + selection.Columns.Count - 1, usedRange.Column + usedRange.Columns.Count - 1)

for (let i = rowFrom; i <= rowTo; i++) {
  const row = API.ActiveSheet.Rows(i) // Determine row
  for (let j = colFrom; j <= colTo; j++) {
    const rg = row.Columns(j) // Specify column from row object to determine cell
    // Text is the display value of the cell, i.e., if the cell was originally a formula, Text is the result after formula calculation
    // If you want to get the uncalculated result, use rg.Formula
    const text = rg.Text
    if (text !== '') {
      rg.Value2 = text // The original Value2 was a formula, clear the formula by rewriting Value2 with the calculated result
    }
  }
}
```

## Highlight Invalid Phone Numbers

Manually select one or more cells, then execute the script. The script will automatically check whether the phone numbers in the cells are valid, and highlight cells with invalid phone numbers after execution.

```js
// Regular expression for phone numbers, matching 11-character strings where the first digit is 1, second digit is 3-9, and remaining digits are numbers
const PhoneReg = /^1[3-9](\d{9})$/i

// Highlight color is yellow
const color = RGB(255, 255, 0)

// Simplified API reference
const API = Application

// User selected area
const selection = API.Selection

// Active range of the workbook
const usedRange = API.ActiveSheet.UsedRange

// Get the minimum intersection of both, i.e., the minimum set to traverse
const rowFrom = Math.max(selection.Row, usedRange.Row)
const rowTo = Math.min(selection.Row + selection.Rows.Count - 1, usedRange.Row + usedRange.Rows.Count - 1)
const colFrom = Math.max(selection.Column, usedRange.Column)
const colTo = Math.min(selection.Column + selection.Columns.Count - 1, usedRange.Column + usedRange.Columns.Count - 1)

for (let i = rowFrom; i <= rowTo; i++) {
    const row = API.ActiveSheet.Rows(i) // Determine row object
    for (let j = colFrom; j <= colTo; j++) {
        const rg = row.Columns(j) // Then determine column object, i.e., determine cell
        const text = rg.Text  // Get the text content of the cell
        if (text !== '' && !PhoneReg.test(text)) { // If regex match fails
            rg.Interior.Color = color // Highlight the cell in yellow
        }
    }
}
```

## Unmerge Cells and Fill with Same Content

Manually select one or more cells, then execute the script. The script will automatically unmerge cells and fill them with the same content.

```js
const API = Application
// User selected area
const selection = API.Selection

// If there are merged cells in the selection
if (selection.MergeCells) {
  // Get all merged cell areas
  const areas = selection.MergeArea.Areas
  const areaCount = areas.Count

  // Iterate through each merged cell area
  for (let areaIndex = 1; areaIndex <= areaCount; areaIndex++) {
    const area = areas.Item(areaIndex)
    const text = area.Formula // Save the content of the merged cell before unmerging

    // Get the range of the merged cell area
    const rowFrom = area.Row
    const rowTo = rowFrom + area.Rows.Count - 1
    const colFrom = area.Column
    const colTo = colFrom + area.Columns.Count - 1

    // Unmerge
    area.UnMerge()
    // Rewrite the content of the original merged cell area with the merged cell value
    for (let i = rowFrom; i <= rowTo; i++) {
      const row = API.ActiveSheet.Rows(i) // Determine row
      for (let j = colFrom; j <= colTo; j++) {
        const cell = row.Columns(j) // Determine column
        cell.Value2 = text // Write the previously saved content to the cell
      }
    }
  }
}
```

## Count Duplicate Occurrences

Manually select one or more cells, then execute the script. The script will automatically calculate the duplicate values and their occurrence counts in the selection, create a new worksheet, and write the duplicate values and counts to the `A2 cell` and `B2 cell` of the new worksheet.

```js
const API = Application
// User selected area
const selection = API.Selection
// Active range of the workbook
const usedRange = API.ActiveSheet.UsedRange

// Get the minimum intersection of both, i.e., the minimum set to traverse
const rowFrom = Math.max(selection.Row, usedRange.Row)
const rowTo = Math.min(selection.Row + selection.Rows.Count - 1, usedRange.Row + usedRange.Rows.Count - 1)
const colFrom = Math.max(selection.Column, usedRange.Column)
const colTo = Math.min(selection.Column + selection.Columns.Count - 1, usedRange.Column + usedRange.Columns.Count - 1)

// countMap is used to count occurrences of each string
const countMap = {}
for (let i = rowFrom; i <= rowTo; i++) {
    const row = API.ActiveSheet.Rows(i) // Determine row
    for (let j = colFrom; j <= colTo; j++) {
        const rg = row.Columns(j) // Then determine column object, i.e., determine cell
        const text = rg.Text // Get the cell value for counting
        if (text) {
            if (countMap[text]) {
                countMap[text]++
            } else {
                countMap[text] = 1
            }
        }
    }
}

const Name = API.ActiveSheet.Name
// Create a new worksheet after the current worksheet with default name
const newSheet = API.Sheets.Add(null, API.Sheets(Name), 1)

// Write 'Duplicate Value' and 'Occurrence Count' to A1 and B1 of the new worksheet
newSheet.Range('A1').Value2 = 'Duplicate Value'
newSheet.Range('B1').Value2 = 'Occurrence Count'

// Starting from the second row, output the string in the first column and its occurrence count in the second column
let index = 2
for (let k in countMap) {
    const row = newSheet.Rows(index)
    row.Columns(1).Value2 = "'" + k
    row.Columns(2).Value2 = countMap[k] + ''
    index++;
}
```
