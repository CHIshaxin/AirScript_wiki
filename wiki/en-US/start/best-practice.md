---
title: "Best Practices"
source: https://airsheet.wps.cn/docs/en-US/start/best-practice.html
updated: 1789700916000
hash: ZUIIzEZC
---

# Best Practices

## User Experience Optimization

#### Run Scripts via Hyperlinks

Scripts can be triggered via hyperlinks. The binding method is as follows:

1. Open **Insert** - **Link** to open the link configuration page.
2. Enter appropriate prompt text in the text input box to describe the script's functionality.
3. Click the dropdown next to **Type**, select **AirScript Script**, and specify the script to bind.
4. Click **OK**.

Click the inserted hyperlink to run the script.

This approach provides a WYSIWYG experience, which is much more user-friendly than opening the script editor and clicking run.

#### More Intuitive Execution Experience

When scripts need to select certain cell ranges, there are multiple implementation methods. Reading and operating user-selected cells through `Application.Selection` provides the most intuitive user experience.

Below is an example of defining traversal range through script variables. This example requires modifying 4 variables each time the range changes, which is very inconvenient and provides a poor experience.

```javascript
// Bad example
const startRow = 1
const startColumn = 1
const endRow = 3
const endColumn = 3
for(let i = startRow; i <= endRow; i++){
  let row=Application.Rows(i)
  for (let j = startColumn; j <= endColumn; j++){
    console.log(row.Columns(j).Text)
  }
}
```

Below is using `Application.Selection` to get the cell range selected by the user in the document interface, which is convenient and intuitive.

```javascript
// Good example
let selection = Application.Selection
let startRow = selection.Row
let startCol = selection.Column
let endRow = startRow + selection.Rows.Count - 1
let endCol = startCol + selection.Columns.Count - 1

for (let i = startRow; i <= endRow; i++) {
    let row = Application.Rows(i)
    for (let j = startCol; j <= endCol; j++) {
        console.log(row.Columns(j).Text)
    }
}
```

## Performance Optimization

#### Reuse Objects as Much as Possible

Each function call to Application uses the script engine to operate on file data, and repeated calls cause performance waste.

For example, when iterating and printing the content of 100*100 cells, reusing objects can double the performance.

Without object reuse:

```javascript
// Bad example
let start = new Date()
for(let i= 1;i<=100;i++){
  for (let j =1;j<=100;j++){
    Application.Rows(i).Columns(j).Text
  }
}
console.log(new Date()-start,'ms')
```

Better approach:

```javascript
// Good example
let start = new Date()
for(let i= 1;i<=100;i++){
  let row=Application.Rows(i)
  for (let j =1;j<=100;j++){
    row.Columns(j).Text
  }
}
console.log(new Date()-start)
```

#### Use UsedRange to Narrow Traversal Range

In the [More Intuitive Execution Experience](#more-intuitive-execution-experience) example above, selecting the entire spreadsheet would result in trillions of iterations, and even selecting just one entire row requires millions of iterations. However, the actual range used in our spreadsheet is not that large. Using `Application.ActiveSheet.UsedRange` can determine the used range of the workbook, thus greatly speeding up script execution time.

Below is an improved version of that example:

```javascript
const selection = Application.Selection
const usedRange = Application.ActiveSheet.UsedRange

// Determine the traversal range
const rowFrom = Math.max(selection.Row, usedRange.Row)
const rowTo = Math.min(selection.Row + selection.Rows.Count - 1, usedRange.Row + usedRange.Rows.Count - 1)
const colFrom = Math.max(selection.Column, usedRange.Column)
const colTo = Math.min(selection.Column + selection.Columns.Count - 1, usedRange.Column + usedRange.Columns.Count - 1)

for (let i = rowFrom; i <= rowTo; i++) {
    const row = Application.ActiveSheet.Rows(i) // Reuse object
    for (let j = colFrom; j <= colTo; j++) {
        console.log(row.Columns(j).Text)
    }
}
```
