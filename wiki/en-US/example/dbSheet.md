---
title: "Multidimensional Table Examples"
source: https://airsheet.wps.cn/docs/en-US/example/dbSheet.html
updated: 1789700916000
hash: B9tby2y_
---

# Multidimensional Table Examples

Here are some example scripts for multidimensional table development, hoping to help you quickly understand and get started with multidimensional table script development.

## Quick Batch Fill Values in Selected Area

```javascript
function main() {
  var activeView = Application.Selection.GetActiveView()

  var selectedRecords = Application.Selection.GetSelectionRecords()[0]
  var time = getNowTime()
  var date = getNowDate()

  // Note: "Date", "Time", "Category" here need to be replaced with the corresponding field names in your table
  Application.Record.UpdateRecords({
    SheetId: activeView.sheetId,
    Records: selectedRecords.map(item => ({
      id: item.id,
      fields: {
        Date: date,
        Time: time,
        Category: 'B'
      }
    }))
  })
}
// Get current time in format "hh:mm:ss"
function getNowTime() {
  return new Date().toTimeString().split(' ')[0]
}
// Get current date in format "yyyy:MM:dd"
function getNowDate() {
  var date = new Date()
  return date.getFullYear() + '/' + (date.getMonth() + 1) + '/' + date.getDate()
}

main()
```

## Quick Implementation of "One-Click Archive"

The following code implements two tables with the same data structure in one file, inserting completed data from table one into table two, and deleting the data from table one.

```javascript
function main() {
  var sheets = Application.Sheet.GetSheets()
  // Filter data that is checked and in category B
  var finishedRecords = Application.Record.GetRecords({
    SheetId: sheets[0].id,
    Filter: {
      mode: 'AND',
      criteria: [
        {
          field: 'Completed',
          op: 'Equals',
          values: ['1']
        },
        {
          field: 'Category',
          op: 'Equals',
          values: ['B']
        }
      ]
    }
  })
  // If filtered data exists
  if (finishedRecords) {
    // Insert data into table two
    Application.Record.CreateRecords({
      SheetId: sheets[1].id,
      Records: finishedRecords.records.map(item => ({
        fields: item.fields
      }))
    })
    // Delete data from table one
    Application.Record.DeleteRecords({
      SheetId: sheets[0].id,
      RecordIds: finishedRecords.records.map(item => item.id)
    })
  }
}
main()
```

Note

Combining the above two examples, you can implement features like automatically setting archive date and time, or one-click archiving of selected records, etc.

## Quickly Create a Table

```javascript
function main() {
  Application.Sheet.CreateSheet({
    Name: 'My Table',
    Fields: [
      { name: 'Name', type: 'MultiLineText' },
      { name: 'Quantity', type: 'Number' },
      { name: 'Date', type: 'Date' },
      { name: 'Time', type: 'Time' },
      { name: 'Checkbox', type: 'Checkbox' },
      { name: 'Hyperlink', type: 'Url' },
      { name: 'Rating', type: 'Rating', max: 5 },
      { name: 'Phone', type: 'Phone' },
      { name: 'ID Card', type: 'ID' },
      { name: 'Currency', type: 'Currency' },
      { name: 'Percentage', type: 'Percentage' },
      { name: 'Email', type: 'Email' },
      { name: 'Progress', type: 'Complete' },
      {
        name: 'Category',
        type: 'SingleSelect',
        items: [{ value: 'A' }, { value: 'B' }, { value: 'C' }]
      },
      {
        name: 'Status',
        type: 'MultipleSelect',
        items: [{ value: 'Completed' }, { value: 'Not Started' }, { value: 'In Progress' }]
      },
      {
        name: 'Contact',
        type: 'Contact',
        multipleContacts: false,
        noticeNewContact: false
      },
      { name: 'Rich Text', type: 'Note' },
      { name: 'Attachment', type: 'Attachment' },
      { name: 'Formula', type: 'Formula' },
      { name: 'Created Time', type: 'CreatedTime' },
      { name: 'Created By', type: 'CreatedBy' },
      { name: 'AutoNumber', type: 'AutoNumber' }
    ],
    Views: [
      { name: 'Grid', type: 'Grid' },
      { name: 'Kanban', type: 'Kanban' },
      { name: 'Gallery', type: 'Gallery' },
      { name: 'Form', type: 'Form' },
      { name: 'Gantt', type: 'Gantt' }
    ]
  })
}
main()
```

## Formatted Data Batch Insert

```javascript
function main() {
  var template = ['Product', 10]
  var records = []
  // Insert 100 category A products
  for (let i = 1; i < 100; i++) {
    records.push({
      Name: template[0] + i,
      Quantity: template[1],
      Category: 'A'
    })
  }
  for (let i = 1; i < 100; i++) {
    records.push({
      Name: template[0] + (100 + i),
      Quantity: template[1] + 10,
      Category: 'B'
    })
  }
  for (let i = 1; i < 100; i++) {
    records.push({
      Name: template[0] + (200 + i),
      Quantity: template[1] + 10,
      Category: 'C'
    })
  }
  var sheet = Application.Selection.GetActiveSheet()
  Application.Record.CreateRecords({
    SheetId: sheet.sheetId,
    Records: records.map(item => ({
      fields: item
    }))
  })
}
main()
```
