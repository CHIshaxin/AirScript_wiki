---
title: "Visible 属性"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/Worksheet/%E5%B1%9E%E6%80%A7/Visible%20%E5%B1%9E%E6%80%A7.html
---
# Visible (属性)

返回或设置一个 XlSheetVisibility 值，它确定对象是否可见。

## 示例

```javascript
/*本示例隐藏 Sheet1。*/
function test() {
    Application.Worksheets.Item("Sheet1").Visible = false
}
```

```javascript
/*本示例将 Sheet1 设置为可见。*/
function test() {
    Application.Worksheets.Item("Sheet1").Visible = true
}
```

```javascript
/*本示例使活动工作簿中的每一张工作表可见。*/
function test() {
    for (let i = 1; i <= Sheets.Count; i++) {
        Application.Sheets.Item(i).Visible = true
    }
}
```

```javascript
/*本示例新建一张工作表，然后将其 Visible 属性设为 xlSheetVeryHidden 要引用该工作表，可使用其对象变量 newSheet，如本示例最后一行所示。*/
function test() {
    let newSheet = Application.Worksheets.Add()
    newSheet.Visible = xlSheetVeryHidden
    newSheet.Range("A1:D4").Formula = "=RAND()"
}
```
