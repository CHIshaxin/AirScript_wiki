---
title: "WorksheetFunction 属性"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/Application/%E5%B1%9E%E6%80%A7/WorksheetFunction%20%E5%B1%9E%E6%80%A7.html
---
# WorksheetFunction (属性)

返回 **[WorksheetFunction](./../../WorksheetFunction/WorksheetFunction 对象.html)** 对象。只读。

## 示例

```javascript
/*本示例对单元格区域 A1:A10 应用 Min 工作表函数，并显示结果。*/
function test() {
    let myRange = Application.Worksheets.Item("Sheet1").Range("A1:A10")
    let answer = Application.WorksheetFunction.Min(myRange)
    console.log(answer)
}
```

```javascript
/*本示例使用 CountA 工作表函数确定列 A 中包含值的单元格数量， 并显示出来。*/
function test() {
    Application.ActiveSheet.Range('A1').Value2 = 123
    Application.ActiveSheet.Range('A3').Value2 = 456
    let colA = Application.Worksheets.Item("sheet1").Columns('A')
    let count = WorksheetFunction.CountA(colA)
    console.log(count)
}
```
