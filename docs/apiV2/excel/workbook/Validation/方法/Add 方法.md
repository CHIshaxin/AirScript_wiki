---
title: "Add 方法"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/Validation/%E6%96%B9%E6%B3%95/Add%20%E6%96%B9%E6%B3%95.html
---
# Add (方法)

向指定区域内添加数据有效性验证。

## 说明

**Add**方法所要求的参数依有效性验证的类型而定，如下表所示。

| 有效性验证类型 | 参数 |
| --- | --- |
| **xlValidateCustom** | **Formula1** 必需，忽略 **Formula2**。**Formula1**必须包含一个表达式，数据项有效时该表达式的值为 **True**，数据项无效时，该值为 **False**。 |
| **xlInputOnly** | 使用 **AlertStyle**、**Formula1** 或 **Formula2**。 |
| **xlValidateList** | **Formula1** 必需，忽略 **Formula2**。**Formula1**必须包含以逗号分隔的值列表，或对该列表的工作表引用。 |
| **xlValidateWholeNumber**、**xlValidateDate**、**xlValidateDecimal**、**xlValidateTextLength** 或 **xlValidateTime** | 必须指定 **Formula1** 或 **Formula2** 之一，或两者均指定。 |

## 参数

| 属性 | 数据类型 | 必填 | 说明 |
| --- | --- | --- | --- |
| Type | XlDVType | 必填 | 有效性验证类型。 |
| AlertStyle | any | 可选 | 有效性验证警告的样式。可为以下 XlDVAlertStyle 常量之一：xlValidAlertInformation、xlValidAlertStop 或 xlValidAlertWarning。 |
| Operator | any | 可选 | 数据有效性验证运算符。可为以下 XlFormatConditionOperator 常量之一：xlBetween、xlEqual、xlGreater、xlGreaterEqual、xlLess、xlLessEqual、xlNotBetween 或 xlNotEqual。 |
| Formula1 | any | 可选 | 数据有效性验证等式中的第一部分。 |
| Formula2 | any | 可选 | 当 Operator 为 xlBetween 或 xlNotBetween 时，数据有效性验证等式的第二部分（其他情况下，此参数被忽略）。 |

## 示例

```javascript
/*本示例为 E5 单元格添加数据有效性验证。*/
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
/*本示例为活动工作表上 B2 单元格添加数据有效性验证。*/
function test() {
    ActiveSheet.Range("B2").Validation.Add(xlValidateDate, xlValidAlertWarning, xlLess, "2024/7/2")
}
```
