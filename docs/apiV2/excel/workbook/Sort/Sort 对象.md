---
title: "Sort 对象"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/Sort/Sort%20%E5%AF%B9%E8%B1%A1.html
---
# Sort (对象)

代表数据区域的排序方式。

## 示例

```javascript
/*此示例为活动工作表区域A1:A11赋值，并对这些数据排序。*/
function test() {
    //Building data to sort on the active sheet.
    Range("A1").Value2 = "Name"
    Range("A2").Value2 = "Bill"
    Range("A3").Value2 = "Rod"
    Range("A4").Value2 = "John"
    Range("A5").Value2 = "Paddy"
    Range("A6").Value2 = "Kelly"
    Range("A7").Value2 = "William"
    Range("A8").Value2 = "Janet"
    Range("A9").Value2 = "Florence"
    Range("A10").Value2 = "Albert"
    Range("A11").Value2 = "Mary"
    console.log("The list is out of order.  Hit Ok to continue...")

    //Selecting a cell within the range.
    Range("A2").Select()

    //Applying sort.
    let sort = Application.ActiveWorkbook.Worksheets.Item(ActiveSheet.Name).Sort
    sort.SortFields.Clear()
    sort.SortFields.Add(Range("A2:A11"), xlSortOnValues, xlAscending, xlSortNormal)
    sort.SetRange(Range("A1:A11"))
    sort.Header = xlYes
    sort.MatchCase = false
    sort.Orientation = xlTopToBottom
    sort.SortMethod = xlPinYin
    sort.Apply()
    console.log("Sort complete.")
}
```

```javascript
/*此示例对第一张工作表区域A1:C1按字符的汉语拼音顺序排序。*/
function test() {
    let sort = Application.Sheets.Item(1).Sort
    sort.SortFields.Clear()
    sort.SortFields.Add(Range("A1:C1"))
    sort.SetRange(Range("A1:C1"))
    sort.Header = xlYes
    sort.MatchCase = false
    sort.Orientation = xlSortRows
    sort.SortMethod = xlPinYin
    sort.Apply()
}
```
