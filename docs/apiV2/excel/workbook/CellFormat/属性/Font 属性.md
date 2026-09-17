---
title: "Font 属性"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/CellFormat/%E5%B1%9E%E6%80%A7/Font%20%E5%B1%9E%E6%80%A7.html
---
# Font (属性)

返回一个 **[Font](./../../Font/Font 对象.html)** 对象，该对象允许用户根据单元格的字体格式设置或返回搜索条件。

## 示例

```javascript
/*此示例设置搜索条件以识别包含红色字体的单元格，应用该条件创建一个单元格，找到该单元格，并通知用户。*/
function test() {
    // Set the search criteria for the font of the cell format.
    Application.FindFormat.Font.ColorIndex = 3

    // Set the color index of the font for cell A5 to red.
    Application.Range("A5").Font.ColorIndex = 3
    Application.Range("A5").Formula = "Red font"
    Application.Range("A1").Select()
    console.log("Cell A5 has red font")

    // Find the cells based on the search criteria.
    Application.Cells.Find("", ActiveCell, xlFormulas, xlPart, xlByRows, xlNext, false, null, true).Activate()
    console.log("ET has found this cell matching the search criteria.")
}
```

```javascript
/*此示例显示查找条件中单元格的字体和字号。*/
function test() {
    let font = Application.FindFormat.Font
    console.log(`字体为${font.Name}，字号为${font.Size}`)
}
```
