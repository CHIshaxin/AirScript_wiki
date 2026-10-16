---
title: "Validation"
source: https://airsheet.wps.cn/docs/ja-JP/api/excel/workbook/Validation.html
updated: 1789700916000
hash: Dj53UTDF
---

# Validation

ワークシート上の範囲に対するデータの入力規則を表します。

Validation オブジェクトのプロパティとメソッドは、以下の一覧を参照してください。

### メソッド

| メソッド | 戻り値の型 | 説明 |
|---|---|---|
| [Add()](#add) | undefined | データの入力規則を追加します |
| [Modify()](#modify) | undefined | データの入力規則を変更します |
| [Delete()](#delete) | undefined | データの入力規則を削除します |

## Add()

データの入力規則を追加します。

### パラメーター

| パラメーター | 型 | 既定値 | 必須 | 説明 |
|---|---|---|---|---|
| Type | Enum |  | はい | 検証の種類。[Enum.XlDVType](/docs/ja-JP/api/excel/workbook/Enum.html#xldvtype) を参照 |
| AlertStyle | Enum |  | はい | 検証メッセージのアイコン。[Enum.XlDVAlertStyle](/docs/ja-JP/api/excel/workbook/Enum.html#xldvalertstyle) を参照 |
| Operator | Enum |  | いいえ | 値または数式の比較演算子。[Enum.XlFormatConditionOperator](/docs/ja-JP/api/excel/workbook/Enum.html#xlformatconditionoperator) を参照 |
| Formula1 | String |  | いいえ | 検証数式の第 1 部分。最大 255 文字 |
| Formula2 | String |  | いいえ | Operator が xlBetween または xlNotBetween のときの第 2 部分（それ以外では無視） |

### 例

```js
// Validation オブジェクト
let validation = Application.Range('A1').Validation

// 入力規則を追加：整数が 1～5 に含まれない場合は警告
validation.Add({
  Type: Application.Enum.XlDVType.xlValidateWholeNumber,
  AlertStyle: Application.Enum.XlDVAlertStyle.xlValidAlertWarning,
  Operator: Application.Enum.XlFormatConditionOperator.xlBetween,
  Formula1: '1',
  Formula2: '5'
})
```

## Modify()

データの入力規則を変更します。

### パラメーター

| パラメーター | 型 | 既定値 | 必須 | 説明 |
|---|---|---|---|---|
| Type | Enum |  | はい | 検証の種類。[Enum.XlDVType](/docs/ja-JP/api/excel/workbook/Enum.html#xldvtype) を参照 |
| AlertStyle | Enum |  | はい | 検証メッセージのアイコン。[Enum.XlDVAlertStyle](/docs/ja-JP/api/excel/workbook/Enum.html#xldvalertstyle) を参照 |
| Operator | Enum |  | いいえ | 比較演算子。[Enum.XlFormatConditionOperator](/docs/ja-JP/api/excel/workbook/Enum.html#xlformatconditionoperator) を参照 |
| Formula1 | String |  | いいえ | 検証数式の第 1 部分。最大 255 文字 |
| Formula2 | String |  | いいえ | Operator が xlBetween または xlNotBetween のときの第 2 部分（それ以外では無視） |

### 例

```js
// Validation オブジェクト
let validation = Application.Range('A1').Validation
// 入力規則を変更
validation.Modify({
  Type: Application.Enum.XlDVType.xlValidateWholeNumber,
  AlertStyle: Application.Enum.XlDVAlertStyle.xlValidAlertWarning,
  Operator: Application.Enum.XlFormatConditionOperator.xlNotBetween,
  Formula1: '23',
  Formula2: '105'
})
```

## Delete()

データの入力規則を削除します。

### 例

```js
// Validation オブジェクト
let validation = Application.Range('A1').Validation

validation.Delete()
```
