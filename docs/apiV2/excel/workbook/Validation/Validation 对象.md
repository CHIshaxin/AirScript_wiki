---
title: "Validation 对象"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/Validation/Validation%20%E5%AF%B9%E8%B1%A1.html
---
# Validation (对象)

代表工作表区域的数据有效性规则。

## 说明

使用 **Validation** 属性可返回 **[Validation](./../Validation/Validation 对象.html)** 对象。

## 示例

```javascript
/*本示例更改单元格 E5 的数据有效性验证。*/
function test() {
    Application.Range("E5").Validation.Modify(xlValidateList, xlValidAlertStop, null, "=$A$1:$A$10")
}
```

使用 **Add**方法可将数据有效性添加到某个区域并创建一个新的 **[Validation](./../Validation/Validation 对象.html)** 对象。

```javascript
/*本示例为 E5 单元格添加数据有效性验证*/
function test() {
    let validation = Application.Range("E5").Validation
    validation.Add(xlValidateWholeNumber, xlValidAlertStop, xlBetween, "5", "10")
    validation.InputTitle = "Integers"
    validation.ErrorTitle = "Integers"
    validation.InputMessage = "Enter an integer from five to ten"
    validation.ErrorMessage = "You must enter a number from five to ten"
}
```
