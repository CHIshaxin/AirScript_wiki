---
title: "数据有效性规则(Validation)"
source: https://airsheet.wps.cn/docs/api/excel/workbook/Validation.html
---
# Validation

代表工作表区域的数据有效性规则

Validation 对象的具体属性和方法请参阅下方的列表。

### 方法列表

| 方法名 | 返回类型 | 简介 |
| --- | --- | --- |
| Add() | undefined | 新增数据有效性规则 |
| Modify() | undefined | 修改数据有效性规则 |
| Delete() | undefined | 删除数据有效性规则 |

## Add()

新增数据有效性规则

### 参数

| 属性 | 数据类型 | 默认值 | 必填 | 说明 |
| --- | --- | --- | --- | --- |
| Type | Enum |  | 是 | 指定要对值进行的有效性测试的类型，可以是 [Enum.XlDVType](Enum.md#xldvtype) 中的值 |
| AlertStyle | Enum |  | 是 | 指定验证过程中显示的消息框所用的图表,可以是[Enum.XlDVAlertStyle](Enum.md#xldvalertstyle) 中的值 |
| Operator | Enum |  | 否 | 指定用于将公式与单元格的值或xlBetween和xlNoteBetween中的值进行比较,比较两个公式的运算符，可以是 [Enum.XlFormatConditionOperator](Enum.md#xlformatconditionoperator) 中的值 |
| Formula1 | String |  | 否 | 数据验证公式中的第一部分，值不得超过 255 个字符 |
| Formula2 | String |  | 否 | 当 Operator 参数为 xlBetween 或 xlNotBetween 时, 数据验证等式的第二部分（否则, 将忽略此参数） |

### 示例

```js
// 数据有效性对象
let validation = Application.Range('A1').Validation

// 添加数据验证,如果输入的值不是整数并且不在1~5之间（包括1和5），则显示警告样式
validation.Add({
  Type: Application.Enum.XlDVType.xlValidateWholeNumber,
  AlertStyle: Application.Enum.XlDVAlertStyle.xlValidAlertWarning,
  Operator: Application.Enum.XlFormatConditionOperator.xlBetween,
  Formula1: '1',
  Formula2: '5'
})
```

## Modify()

修改数据有效性规则

### 参数

| 属性 | 数据类型 | 默认值 | 必填 | 说明 |
| --- | --- | --- | --- | --- |
| Type | Enum |  | 是 | 指定要对值进行的有效性测试的类型，可以是 [Enum.XlDVType](Enum.md#xldvtype) 中的值 |
| AlertStyle | Enum |  | 是 | 指定验证过程中显示的消息框所用的图表,可以是[Enum.XlDVAlertStyle](Enum.md#xldvalertstyle) 中的值 |
| Operator | Enum |  | 否 | 指定用于将公式与单元格的值或xlBetween和xlNoteBetween中的值进行比较,比较两个公式的运算符，可以是 [Enum.XlFormatConditionOperator](Enum.md#xlformatconditionoperator) 中的值 |
| Formula1 | String |  | 否 | 数据验证公式中的第一部分，值不得超过 255 个字符 |
| Formula2 | String |  | 否 | 当 Operator 参数为 xlBetween 或 xlNotBetween 时, 数据验证等式的第二部分（否则, 将忽略此参数） |

### 示例

```js
// 数据有效性对象
let validation = Application.Range('A1').Validation
// 修改数据验证规则
validation.Modify({
  Type: Application.Enum.XlDVType.xlValidateWholeNumber,
  AlertStyle: Application.Enum.XlDVAlertStyle.xlValidAlertWarning,
  Operator: Application.Enum.XlFormatConditionOperator.xlNotBetween,
  Formula1: '23',
  Formula2: '105'
})
```

## Delete()

删除数据有效性规则

### 示例

````js
// 数据有效性对象
let validation = Application.Range('A1').Validation

validation.Delete()
```、
````
