---
title: "ShowInput 属性"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/Validation/%E5%B1%9E%E6%80%A7/ShowInput%20%E5%B1%9E%E6%80%A7.html
---
# ShowInput (属性)

如果用户在数据有效性检查区域内选定了某一单元格时，显示数据有效性检查输入消息，则该属性值为 **True**。**Boolean** 类型，可读写。

## 示例

```javascript
/*本示例向第一张工作表上的单元格 A10 中添加数据有效性检查。输入的值必须处于 5 到 10 之间；如果用户输入了无效数据，将显示错误消息，但不显示输入消息。*/
function test() {
    let validation = Application.Worksheets.Item(1).Range("A10").Validation
    validation.Add(xlValidateWholeNumber, xlValidAlertStop, xlBetween, "5", "10")
    validation.ErrorMessage = "value must be between 5 and 10"
    validation.ShowInput = false
    validation.ShowError = true
}
```

```javascript
/*本示例显示活动工作表上单元格 E1 的数据有效性检验是否展示输入消息。*/
function test() {
    let validation = ActiveSheet.Range("E1").Validation
    console.log(validation.ShowInput)
}
```
