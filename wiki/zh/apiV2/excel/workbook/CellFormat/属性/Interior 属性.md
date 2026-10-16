---
title: "Interior (属性)"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/CellFormat/%E5%B1%9E%E6%80%A7/Interior%20%E5%B1%9E%E6%80%A7.html
updated: 1789700916000
hash: CtjG1w32
---

# Interior (属性)

返回一个 **[Interior](./../../Interior/Interior 对象.html)** 对象，该对象允许用户根据单元格内部格式设置或返回搜索条件。

## 示例

```javascript
/*此示例设置搜索条件以识别内部为纯黄色的单元格，创建满足该条件的单元格，找到该单元格，并通知给用户。*/
function test() {
    // Set the search criteria for the interior of the cell format.
    let interior = Application.FindFormat.Interior
    interior.ColorIndex = 6
    interior.Pattern = xlSolid
    interior.PatternColorIndex = xlAutomatic

    // Create a yellow interior for cell A5.
    Application.Range("A5").Select()
    Application.Selection.Interior.ColorIndex = 6
    Application.Selection.Interior.Pattern = xlSolid
    Application.Selection.Interior.PatternColorIndex = xlAutomatic
    Range("A1").Select()
    console.log("Cell A5 has a yellow interior.")

    // Find the cells based on the search criteria.
    Application.Cells.Find("", ActiveCell, xlFormulas, xlPart, xlByRows, xlNext, false, null, true).Activate()
    console.log("ET has found this cell matching the search criteria.")
}
```

```javascript
/*此示例显示查找条件中单元格的内部颜色是否为黄色。*/
function test() {
    console.log(Application.FindFormat.Interior.Color == RGB(255, 255, 0))
}
```
