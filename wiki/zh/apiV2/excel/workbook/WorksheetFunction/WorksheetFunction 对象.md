---
title: "WorksheetFunction (对象)"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/WorksheetFunction/WorksheetFunction%20%E5%AF%B9%E8%B1%A1.html
updated: 1789700916000
hash: BJiY4oeD
---

# WorksheetFunction (对象)

用作可从 Visual Basic 中调用的 ET 工作表函数的容器。

## 说明

使用 **WorksheetFunction**属性可返回 **[WorksheetFunction](./../WorksheetFunction/WorksheetFunction 对象.html)** 对象。下例显示给区域 A1:A10 应用 **Min** 工作表函数的结果。

```javascript
function test(){
let myRange = Application.Worksheets.Item("Sheet1").Range("A1:C10")
let answer = Application.WorksheetFunction.Min(myRange)
console.log(answer)
}
```
