---
title: "Validation (属性)"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/Range/%E5%B1%9E%E6%80%A7/Validation%20%E5%B1%9E%E6%80%A7.html
updated: 1789700916000
hash: BIJ4Q0K9
---

# Validation (属性)

返回一个 **[Validation](./../../Validation/Validation 对象.html)** 对象，该对象表示指定区域的数据有效性检验。只读。

## 示例

```javascript
/*本示例使单元格 E5 的数据有效性检验为允许有空值。*/
function test() {
    Range("E5").Validation.IgnoreBlank = true
}
```

```javascript
/*本示例为 G5 单元格添加数据有效性验证*/
function test() {
    let validation = Application.Range("G5").Validation
    validation.Add(xlValidateWholeNumber,  xlValidAlertStop, xlBetween, "1", "100")
    validation.InputTitle = "Integers"
    validation.ErrorTitle = "Integers"
    validation.InputMessage = "Enter an integer from 1 to 100"
    validation.ErrorMessage = "You must enter a number from 1 to 100"
}
```
