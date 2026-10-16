---
title: "Validation"
source: https://airsheet.wps.cn/docs/en-US/api/excel/workbook/Validation.html
updated: 1789700916000
hash: HH6MuRA-
---

# Validation

Represents data validation rules for a range on a worksheet.

See the lists below for the Validation object's properties and methods.

### Methods

| Method | Return type | Description |
|---|---|---|
| [Add()](#add) | undefined | Adds a data validation rule |
| [Modify()](#modify) | undefined | Modifies a data validation rule |
| [Delete()](#delete) | undefined | Deletes data validation rules |

## Add()

Adds a data validation rule.

### Parameters

| Property | Data type | Default | Required | Description |
|---|---|---|---|---|
| Type | Enum |  | Yes | Validation type; see [Enum.XlDVType](/docs/en-US/api/excel/workbook/Enum.html#xldvtype) |
| AlertStyle | Enum |  | Yes | Icon for validation messages; see [Enum.XlDVAlertStyle](/docs/en-US/api/excel/workbook/Enum.html#xldvalertstyle) |
| Operator | Enum |  | No | Comparison operator for values or formulas; see [Enum.XlFormatConditionOperator](/docs/en-US/api/excel/workbook/Enum.html#xlformatconditionoperator) |
| Formula1 | String |  | No | First part of the validation formula; max 255 characters |
| Formula2 | String |  | No | Second part when Operator is xlBetween or xlNotBetween (otherwise ignored) |

### Example

```js
// Validation object
let validation = Application.Range('A1').Validation

// Add validation: whole number not in 1-5 shows warning
validation.Add({
  Type: Application.Enum.XlDVType.xlValidateWholeNumber,
  AlertStyle: Application.Enum.XlDVAlertStyle.xlValidAlertWarning,
  Operator: Application.Enum.XlFormatConditionOperator.xlBetween,
  Formula1: '1',
  Formula2: '5'
})
```

## Modify()

Modifies a data validation rule.

### Parameters

| Property | Data type | Default | Required | Description |
|---|---|---|---|---|
| Type | Enum |  | Yes | Validation type; see [Enum.XlDVType](/docs/en-US/api/excel/workbook/Enum.html#xldvtype) |
| AlertStyle | Enum |  | Yes | Icon for validation messages; see [Enum.XlDVAlertStyle](/docs/en-US/api/excel/workbook/Enum.html#xldvalertstyle) |
| Operator | Enum |  | No | Comparison operator; see [Enum.XlFormatConditionOperator](/docs/en-US/api/excel/workbook/Enum.html#xlformatconditionoperator) |
| Formula1 | String |  | No | First part of the validation formula; max 255 characters |
| Formula2 | String |  | No | Second part when Operator is xlBetween or xlNotBetween (otherwise ignored) |

### Example

```js
// Validation object
let validation = Application.Range('A1').Validation
// Modify validation
validation.Modify({
  Type: Application.Enum.XlDVType.xlValidateWholeNumber,
  AlertStyle: Application.Enum.XlDVAlertStyle.xlValidAlertWarning,
  Operator: Application.Enum.XlFormatConditionOperator.xlNotBetween,
  Formula1: '23',
  Formula2: '105'
})
```

## Delete()

Deletes data validation rules.

### Example

```js
// Validation object
let validation = Application.Range('A1').Validation

validation.Delete()
```
