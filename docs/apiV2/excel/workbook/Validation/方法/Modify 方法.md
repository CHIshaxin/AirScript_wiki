---
title: "Modify 方法"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/Validation/%E6%96%B9%E6%B3%95/Modify%20%E6%96%B9%E6%B3%95.html
---
# Modify (方法)

修改指定区域的数据有效性验证。

## 说明

**Modify**方法所要求的参数依有效性验证的类型而定，如下表所示。

| 有效性验证类型 | 参数 |
| --- | --- |
| **xlInputOnly** | 不使用 **AlertStyle**、**Formula1** 和 **Formula2**。 |
| **xlValidateCustom** | **Formula1** 必需，忽略 **Formula2**。**Formula1**必须包含一个表达式，数据项有效时该表达式的值为 **True**，数据项无效时，该值为 **False**。 |
| **xlValidateList** | **Formula1** 必需，忽略 **Formula2**。**Formula1**必须包含一个以逗号分隔的值列表，或对该列表的工作表引用。 |
| **xlValidateDate**、**xlValidateDecimal**、**xlValidateTextLength**、 **xlValidateTime** 或 **xlValidateWholeNumber** | 必须指定 **Formula1** 或 **Formula2**，或两者均指定。 |

## 参数

| 属性 | 数据类型 | 必填 | 说明 |
| --- | --- | --- | --- |
| Type | any | 可选 | 一个代表有效性验证类型的 XlDVType 值。 |
| AlertStyle | any | 可选 | 一个代表有效性验证警告样式的 XlDVAlertStyle 值。 |
| Operator | any | 可选 | 一个代表数据有效性验证运算符的 XlFormatConditionOperator 值。 |
| Formula1 | any | 可选 | 数据有效性验证等式中的第一部分。 |
| Formula2 | any | 可选 | 当 Operator 为 xlBetween 或 xlNotBetween 时，数据有效性的第二部分；其他情况下，此参数被忽略。 |

## 示例

```javascript
/*本示例更改单元格 B1 的数据有效性验证。*/
function test() {
    Application.Range("B1").Validation.Modify(xlValidateList, xlValidAlertStop, xlBetween, "=$A$1:$A$10")
}
```

```javascript
/*本示例更改活动工作表上单元格 B1 的数据有效性验证。*/
function test() {
    let validation = Application.ActiveSheet.Range("B1").Validation
    validation.Modify(xlValidateDate, xlValidAlertWarning, xlNotBetween, "2024/7/1", "2024/7/7")
}
```
