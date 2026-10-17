---
title: "Sheet"
source: https://airsheet.wps.cn/docs/ja-JP/api/excel/databook/Sheet.html
updated: 1789700916000
hash: DP6ZWRgS
---

# Sheet

ブック（`Workbook`）内の 1 枚のデータ Sheet（`Sheet`）オブジェクト。

プロパティとメソッドは以下の一覧を参照してください。

#### プロパティ

| 名前 | 型 | 説明 |
| --- | --- | --- |
| [Id](#id) | String | Sheet ID |
| [Name](#name) | String | Sheet 名 |
| [Index](#index) | Number | すべての Sheet 中のインデックス |
| [Visible](#visible) | Boolean | Sheet が表示されているかどうか |
| [Type](#type) | String | Sheet の種類 |
| [Field](#field) | [Field](#field) | この Sheet 上の Field |
| [Record](#record) | [Record](#record) | この Sheet 上の行 Record |

#### メソッド

| 名前 | 戻り値の型 | 説明 |
| --- | --- | --- |
| [Activate()](#activate) | undefined | Sheet をアクティブにする |
| [Move()](#move) | undefined | Sheet を移動する |
| [Delete()](#delete) | undefined | Sheet を削除する |
| [IsDBSheet()](#isdbsheet) | Boolean | データ Sheet かどうか |

## Id

Sheet ID を取得します。

#### 型

String – Sheet ID

#### 例

```js
const sheet = Application.ActiveSheet
console.log(sheet.Id)
```

## Name

Sheet 名の取得または設定。

#### 型

String – すべての Sheet 間での名前

#### 例

```js
const sheet = Application.ActiveSheet
console.log(sheet.Name) // Sheet2
sheet.Name = 'WPS WebOffice'
```

## Index

すべての Sheet の中でのこの Sheet のインデックス。

#### 型

String – すべての Sheet 中のインデックス（元の API と同様）

#### 例

```js
const sheet = Application.ActiveSheet
console.log(sheet.Index) // 1
```

## Visible

Sheet の表示／非表示。

#### 型

Boolean – 表示状態

#### 例

```js
const sheet = Application.ActiveSheet
sheet.Visible = false
sheet.Visible = true
```

## Type

Sheet の種類。

#### 型

[Enum.xlEtDataBaseSheet](Enum.md#xlsheettype) – Sheet の種類

#### 例

```js
const sheet = Application.ActiveSheet
console.log(sheet.Type) // xlEtDataBaseSheet
```

## Field

データ Sheet 上の Field。[Field](Field.md) オブジェクトを返します。

#### 型

[Field](Field.md)

#### 例

```js
const sheet = Application.ActiveSheet
const fields = sheet.Field.GetFields()
```

## Record

データ Sheet 上の行 Record。[Record](Record.md) オブジェクトを返します。

#### 型

[Record](Record.md)

#### 例

```js
const sheet = Application.ActiveSheet
const record = sheet.Record.GetRecord({ RecordId: 'Bz' })
```

## Activate()

Sheet をアクティブにします。

#### 例

```js
const sheet = Application.Sheets.Item(1)
sheet.Activate()
```

## Move()

Sheet を移動します。

#### パラメーター

**`Before` と `After` は同時に使用できません。**

| プロパティ | 型 | 既定値 | 必須 | 説明 |
| --- | --- | --- | --- | --- |
| Before | number | null | いいえ | この Sheet の前に配置する Sheet の ID。`After` と同時に使用不可。 |
| After | number | null | いいえ | この Sheet の後に配置する Sheet の ID。`Before` と同時に使用不可。 |

#### 例

```js
const sheet = Application.ActiveSheet
sheet.Move({
  Before: null,
  After: Application.Sheets(2).Id
})
```

## Delete()

Sheet を削除します。

#### 戻り値

undefined

#### 例

```js
Application.Sheets.Item('Sheet2').Delete()
```

## IsDBSheet()

データ Sheet かどうか。

#### 戻り値

Boolean

#### 例

```js
Application.ActiveSheet.IsDBSheet()
```
