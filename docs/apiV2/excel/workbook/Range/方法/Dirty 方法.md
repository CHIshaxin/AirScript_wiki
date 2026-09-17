---
title: "Dirty 方法"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/Range/%E6%96%B9%E6%B3%95/Dirty%20%E6%96%B9%E6%B3%95.html
---
# Dirty (方法)

下一次重新计算发生时指定要重新计算的区域。

## 说明

对于 ET 知道需要进行重新计算的单元格，**Calculate** 方法强制重新计算指定的区域。

如果应用程序处于手动计算模式，请使用 **Dirty** 方法使 ET 识别要重新计算的单元格。如果应用程序处于自动计算模式，请使用 **Dirty** 方法使 ET 执行重新计算。

## 示例

```javascript
/*本示例中，Microsoft Excel 在单元格 A3 中输入了一个公式，保存了更改，并重新计算了单元格 A3。*/
function test() {
    console.log("Two values and a formula will be entered.")
    Range("A1").Value(undefined, 1)
    Range("A2").Value(undefined, 2)
    Range("A3").Formula = "=A1+A2"
    // Save the changes made to the worksheet.
    Application.DisplayAlerts = false
    Application.Save()
    console.log("Changes saved.")
    // Force a recalculation of range A3.
    Application.Range("A3").Dirty()
    console.log("Try to close the file without saving and a dialog box will appear.")
}
```

```javascript
/*本示例重新计算 A1:C3 单元格*/
function test() {
    Range("A1:C3").Dirty()
}
```
