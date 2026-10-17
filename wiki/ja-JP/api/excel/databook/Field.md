---
title: "Field"
source: https://airsheet.wps.cn/docs/ja-JP/api/excel/databook/Field.html
updated: 1789700916000
hash: Ccv9wcaW
---

# Field

Field の操作。

### メソッド

| メソッド | 戻り値の型 | 説明 |
| --- | --- | --- |
| [GetFields()](#getfields) | Array | Field 定義を取得 |
| [CreateFields()](#createfields) | Array | Field を作成 |
| [DeleteFields()](#deletefields) | Array | Field を削除 |
| [UpdateFields()](#updatefields) | Array | Field を更新 |

## GetFields()

Field 定義を取得します。

### 戻り値

Array – Sheet 上のすべての Field

| プロパティ | 型 | 説明 |
| --- | --- | --- |
| id | String | Field ID |
| name | String | Field 名 |
| type | String | [Field の種類](Appendix.md#appendix-1-datasheet-field-types) |

### 例

```javascript
const sheet = Application.ActiveSheet
const fields = sheet.Field.GetFields()
console.log(fields)
// [
//  {"id":"Ce","name":"Name","type":"MultiLineText"},
//  {"id":"Cf","name":"Quantity","type":"Number"},
// ]
```

## CreateFields()

Field を作成します。

### パラメーター

| プロパティ | 型 | 既定値 | 必須 | 説明 |
| --- | --- | --- | --- | --- |
| Fields | Array |  | はい | Field 定義。形式は [付録](Appendix.md#appendix-1-datasheet-field-types) を参照 |

### 戻り値

Array – 作成された Field

| プロパティ | 型 | 説明 |
| --- | --- | --- |
| id | String | Field ID |
| name | String | Field 名 |
| type | String | [Field の種類](Appendix.md#appendix-1-datasheet-field-types) |

### 例

```javascript
const sheet = Application.ActiveSheet
const field = sheet.Field.CreateFields({
  Fields: [{ name: 'Score', type: 'Rating', max: 5 }]
})
console.log(field)
// [{"id":"LZ","name":"Score","type":"Rating"}]
```

## DeleteFields()

Field を削除します。

### パラメーター

| プロパティ | 型 | 既定値 | 必須 | 説明 |
| --- | --- | --- | --- | --- |
| Fields | Array |  | はい | 削除する Field の ID |

### 戻り値

Array – Field ごとの削除結果

| プロパティ | 型 | 説明 |
| --- | --- | --- |
| id | String | Field ID |
| deleted | Boolean | 削除に成功したかどうか |

### 例

```javascript
const sheet = Application.ActiveSheet
const resutlt = sheet.Field.DeleteFields({ FieldIds: ['P', 'Q'] })
console.log(resutlt)
// [{"deleted":false,"id":"P"},{"deleted":false,"id":"Q"}]
```

## UpdateFields()

Field を更新します。

### パラメーター

| プロパティ | 型 | 既定値 | 必須 | 説明 |
| --- | --- | --- | --- | --- |
| Fields | Array |  | はい | Field ID と名前を含む更新内容。形式は [付録](Appendix.md#appendix-1-datasheet-field-types) を参照 |

### 戻り値

Array – 更新後の Field 定義

| プロパティ | 型 | 説明 |
| --- | --- | --- |
| id | String | Field ID |
| name | String | Field 名 |
| type | String | [Field の種類](Appendix.md#appendix-1-datasheet-field-types) |

### 例

```javascript
const sheet = Application.ActiveSheet
sheet.Field.UpdateFields({
    Fields: [{ id: 'LG', name: 'Jump' }]
})
```
