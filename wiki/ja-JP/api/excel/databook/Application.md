---
title: "Application"
source: https://airsheet.wps.cn/docs/ja-JP/api/excel/databook/Application.html
updated: 1789700916000
hash: C7RThpWl
---

# Application

ドキュメント操作の最上位オブジェクトです。ドキュメント関連の処理はすべて、このオブジェクトを直接または間接的に通じて行われます。

`Application` はファイルの最上位オブジェクトです。新しいファイルを開いた場合も `Application` が返されます。

スクリプト内では、`Application` は現在のファイルの最上位オブジェクトを指します。常に 1 つだけ存在します。

`Application` のプロパティとメソッドは、以下の一覧を参照してください。

#### プロパティ

| プロパティ | 型 | 説明 |
|---|---|---|
| [ActiveSheet](#activesheet) | [Sheet](/docs/ja-JP/api/excel/databook/Sheet.html) | アクティブなワークシート／データ Sheet |
| [Sheets](#sheets) | [Sheets](/docs/ja-JP/api/excel/databook/Application.html#sheets) | ファイル内のすべてのワークシート／データ Sheet |
| [FileInfo](#fileinfo) | オブジェクト | ドキュメントに関する情報 |
| [UserInfo](#userinfo) | オブジェクト | ドキュメントのユーザー情報 |
| [Enum](#enum) | [Enum](/docs/ja-JP/api/excel/databook/Enum.html) | すべての列挙型 |

#### メソッド

| メソッド | 戻り値の型 | 説明 |
|---|---|---|
| [Sheets(name)](#sheets-1) | [Sheet](/docs/ja-JP/api/excel/databook/Sheet.html) | 名前が `name` のワークシート／データ Sheet を取得 |

## ActiveSheet

アクティブなワークシート／データ Sheet です。`Sheet.Activate()` で切り替えられます。このプロパティは [Sheet](/docs/ja-JP/api/excel/databook/Sheet.html#activesheet) オブジェクトを返すため、アクティブな Sheet を操作できます。

スクリプトはサーバー上で隔離された環境で実行されるため、スクリプト環境の `ActiveSheet` がユーザーの `ActiveSheet` と一致しない場合があります。

ルール:

1. スクリプト実行時、スクリプト環境の `ActiveSheet` は、ユーザーの現在の `ActiveSheet` に合わせて設定されます。
2. スクリプトがスクリプト環境の `ActiveSheet` を変更しても、ユーザーの `ActiveSheet` は変わりません。

#### 型

[Sheet](/docs/ja-JP/api/excel/databook/Sheet.html) – アクティブなワークシート／データ Sheet

#### 例

```js
console.log(Application.ActiveSheet.Name) // DataSheet2

// 名前が DataSheet2 の Sheet に切り替え
Application.Sheets.Item('DataSheet2').Activate()
console.log(Application.ActiveSheet.Name) // DataSheet2
```

## Sheets

現在のファイルで操作可能なすべての `Sheet` を [Sheets](/docs/ja-JP/api/excel/databook/Application.html#sheets) オブジェクトとして返します。

#### 型

[Sheets](/docs/ja-JP/api/excel/databook/Application.html#sheets)

#### 例

```js
// ブック内のすべてのワークシート／データ Sheet（Sheet）のコレクション。以下は同等
let sheets = Application.ActiveWorkbook.Sheets
sheets = Application.Sheets

// すべてのワークシート／データ Sheet の名前を出力
for (let i = 1; i <= sheets.Count; i++) {
  console.log(sheets.Item(i).Name)
}
```

### Sheets.Count

ワークシート／データ Sheet の数。

#### 型

Number – ブック内のワークシート／データ Sheet の件数

#### 例

```js
// 以下は同等
let sheets = Application.ActiveWorkbook.Sheets
sheets = Application.Sheets

console.log(sheets.Count) // 1
```

### Sheets.DefaultNewSheetName

新しいワークシートの既定の名前。

#### 戻り値の型

String – Sheet を追加するときに名前を指定しない場合に使用できる名前

#### 例

```js
const defaultName = Application.Sheets.DefaultNewSheetName
Application.Sheets.Add(
  null,
  Application.ActiveSheet.Name,
  1,
  Application.Enum.XlSheetType.xlWorksheet,
  defaultName
)
```

### Sheets.Add()

ワークシートを追加します。`Before` と `After` の両方が指定された場合は、`Before` が優先されます。

#### パラメーター

| プロパティ | 型 | 既定値 | 必須 | 説明 |
|---|---|---|---|---|
| Before | String/Number |  | いいえ | `After` が空のとき必須。既存 Sheet のインデックスまたは名前。新しい Sheet はその前に挿入 |
| After | String/Number |  | いいえ | `Before` が空のとき必須。既存 Sheet のインデックスまたは名前。新しい Sheet はその後に挿入 |
| Count | Number | 1 | いいえ | 追加する Sheet の数。既定は選択中の Sheet の数 |
| Type | Enum |  | いいえ | Sheet の種類。[Enum.XlSheetType](/docs/ja-JP/api/excel/databook/Enum.html#xlsheettype) を参照 |
| Name | Name |  | いいえ | 新しい Sheet の名前 |

#### 例

```js
Application.Sheets.Add(
  null,
  Application.ActiveSheet.Name,
  1,
  Application.Enum.XlSheetType.xlWorksheet,
  'NewSheet'
)
```

### Sheets.Item()

名前またはインデックスで Sheet を選択します。

#### パラメーター

| プロパティ | 型 | 既定値 | 必須 | 説明 |
|---|---|---|---|---|
| index | String/Number |  | はい | Sheet 名またはインデックス |

#### 戻り値の型

[Sheet](/docs/ja-JP/api/excel/databook/Sheet.html) – 一致したワークシート／データ Sheet

#### 例

```js
Application.Sheets.Item('Sheet2').Activate()
Application.Sheets.Item(1).Activate()
```

### Sheets.Each()

すべての Sheet を反復し、コールバックを実行します。

#### パラメーター

| プロパティ | 型 | 既定値 | 必須 | 説明 |
|---|---|---|---|---|
| callback | Function | null | はい | `Array.prototype.forEach` に似た動作 |

#### 例

```js
Application.Sheets.Each(function (item) {
  console.log(item.Name) // Sheet1 Sheet2
})
```

## FileInfo

現在のファイルに関する基本情報。

#### 型

オブジェクト – ファイルのメタデータ

| 名前 | 型 | 説明 |
|---|---|---|
| id | string | ファイル ID |
| name | string | ファイル名 |
| officeType | string | ドキュメントの種類 |
| creator | [CreatorObject](#creatorobject) | 作成者情報 |
| size | number | ファイルサイズ |
| groupId | string | グループ ID |
| docType | number | ドキュメントの種類（数値） |

#### CreatorObject

| 名前 | 型 | 説明 |
|---|---|---|
| id | string | 作成者 ID |
| name | string | 作成者名 |
| avatar_url | string | アバター URL |
| logined | boolean | ログイン済みかどうか |
| attrs | オブジェクト | プロパティ |
| real_id | string | 実 ID |

#### 例

```javascript
console.log(Application.FileInfo)
/*{
 "id": "<open_id>",
 ...
}*/
```

## UserInfo

現在のファイルのユーザー情報。

#### 型

オブジェクト – ユーザーのメタデータ

| 名前 | 型 | 説明 |
|---|---|---|
| id | string | ユーザー ID |
| name | string | ユーザー名 |

#### 例

```javascript
console.log(Application.UserInfo)
```

## Enum

列挙型。`Application` から公開されます。

`Application.Enum` を使用します。

#### 型

[Enum](/docs/ja-JP/api/excel/databook/Enum.html) – すべての列挙型

#### 例

```js
console.log(Application.Enum.XlSheetType)
// {"xlChart":-4109,"xlDialogSheet":-4116,"xlExcel4IntlMacroSheet":4,"xlExcel4MacroSheet":3,"xlWorksheet":-4167}
```

## Sheets()

`Sheets.Item()` の代わりに関数として使用します。[Sheet](/docs/ja-JP/api/excel/databook/Sheet.html) を返します。

#### パラメーター

| パラメーター | 型 | 必須 | 説明 |
|---|---|---|---|
| name | string | はい | ワークシート／データ Sheet の名前 |

#### 戻り値の型

[Sheet](/docs/ja-JP/api/excel/databook/Sheet.html) – その名前の Sheet オブジェクト

#### 例

```js
console.log(Application.Sheets.Count) // 1

console.log(Application.Sheets('Sheet2').Range('A1').Text)
console.log(Application.Sheets.Item('Sheet2').Range('A1').Text)
```
