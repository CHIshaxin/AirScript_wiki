---
title: "ErrorTitle 属性"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/Validation/%E5%B1%9E%E6%80%A7/ErrorTitle%20%E5%B1%9E%E6%80%A7.html
---
# ErrorTitle (属性)

返回或设置数据有效性错误对话框的标题。**String** 类型，可读写。

## 示例

```javascript
/*本示例向单元格 E5 添加数据有效性验证，并指定输入信息和错误消息。*/
function test() {
    let validation = Application.Range("E5").Validation
    validation.Add(xlValidateWholeNumber, xlValidAlertStop, xlBetween, "5", "10")
    validation.InputTitle = "Integers"
    validation.ErrorTitle = "Integers"
    validation.InputMessage = "Enter an integer from five to ten"
    validation.ErrorMessage = "You must enter a number from five to ten"
}
```

```javascript
/*本示例显示活动工作表上单元格 B1 的数据有效性验证的错误对话框的标题。*/
function test() {
    let validation = ActiveSheet.Range("B1").Validation
    console.log(`标题为：${validation.ErrorTitle}`)
}
```
