---
title: "Record"
source: https://airsheet.wps.cn/docs/ja-JP/api/excel/databook/Record.html
updated: 1789700916000
hash: Fab9nyDQ
---

# Record

行の Record。

### メソッド

| メソッド | 戻り値の型 | 説明 |
| --- | --- | --- |
| [GetRecords()](#getrecords) | Array | 複数行を取得 |
| [GetRecord()](#getrecord) | Object | 1 行を取得 |
| [DeleteRecords()](#deleterecords) | Array | 行を削除 |
| [UpdateRecords()](#updaterecords) | Array | 行を更新 |
| [CreateRecords()](#createrecords) | Array | 行を作成 |
| [GetAttachmentURL()](#getattachmenturl) | String | アップロード済みの添付ファイルまたは画像の URL |

## GetRecords()

複数行の Record を取得します。

注意

1 リクエストあたり最大 100 行まで返します。大きなデータセットではページングを使用してください。

### パラメーター

| プロパティ | 型 | 既定値 | 必須 | 説明 |
| --- | --- | --- | --- | --- |
| ViewId | String |  | いいえ | 指定した場合、ユーザーが View で見ているのと同じ状態でその View から Record を読み取ります。未指定の場合は Sheet から |
| PageSize | Number | 100 | いいえ | ページング時の本クエリの先頭 Record（含む）。空の場合は先頭の Record から。**現在の上限: 1000** |
| Offset | Number |  | いいえ | ページング用。レスポンスに次ページ先頭行の `offset` が含まれる場合があります。最終ページ、または `maxRecords` に達した場合は省略 |
| MaxRecords | Number |  | いいえ | 一致する行の先頭から最大 `maxRecords` 行のみ返します。省略時は一致するすべての行 |
| Fields | Array |  | いいえ | [Field の種類](Appendix.md#appendix-1-datasheet-field-types) |
| Filter | Object |  | いいえ | [付録 3](Appendix.md#appendix-3-filter-criteria) を参照 |

### 戻り値

オブジェクト – Record のバッチ

| プロパティ | 型 | 説明 |
| --- | --- | --- |
| Offset | String | ページング用カーソル。次のリクエストに渡して続きを取得 |
| Records | Array[Object] | Record |

#### Record の形状

| プロパティ | 型 | 説明 |
| --- | --- | --- |
| id | String | Record ID |
| Fields | Object | Field ID または名前による値。形式は [付録](Appendix.md#appendix-1-datasheet-field-types) を参照 |

### 例

```javascript
const sheet = Application.ActiveSheet
function fetchAllRecords() {
  const view = sheet.Selection.GetActiveView()
  let all = []
  let offset = null

  while (all.length === 0 || offset) {
    let records = sheet.Record.GetRecords({
      ViewId: view.viewId,
      Offset: offset
    })
    offset = records.offset
    all = all.concat(records.records)
  }
  console.log(all.length)
  return all
}

fetchAllRecords()
```

## GetRecord()

1 行を取得します。

### パラメーター

| プロパティ | 型 | 既定値 | 必須 | 説明 |
| --- | --- | --- | --- | --- |
| RecordId | String |  | はい | Record ID |

### 戻り値

オブジェクト – 1 件の Record

| プロパティ | 型 | 説明 |
| --- | --- | --- |
| id | String | Record ID |
| Fields | Object | Field の値。形式は [付録](Appendix.md#appendix-1-datasheet-field-types) を参照 |

### 例

```javascript
const sheet = Application.ActiveSheet
const record = sheet.Record.GetRecord({ RecordId: 'Bz' })
console.log(record)
// {"fields":{"Date":"2023/02/21"},"id":"Bz"}
```

## DeleteRecords()

行を削除します。

### パラメーター

| プロパティ | 型 | 既定値 | 必須 | 説明 |
| --- | --- | --- | --- | --- |
| RecordIds | Array |  | はい | 削除する ID |

### 戻り値

Array – ID ごとの結果

| プロパティ | 型 | 説明 |
| --- | --- | --- |
| id | String | Record ID |
| deleted | Boolean | 削除できた場合は `true`、できなかった場合は `false` |

### 例

```javascript
const sheet = Application.ActiveSheet
const result = sheet.Record.DeleteRecords({
  RecordIds: ['J', 'P', 'Q']
})
console.log(resutlt)
// [{"deleted":true,"id":"P"},{"deleted":false,"id":"Q"}]
```

## UpdateRecords()

行を更新します。

### パラメーター

| プロパティ | 型 | 既定値 | 必須 | 説明 |
| --- | --- | --- | --- | --- |
| Records | Array[Object] |  | はい | 更新する Record |

#### Record エントリ

| プロパティ | 型 | 説明 |
| --- | --- | --- |
| id | String | Record ID |
| Fields | Object | Field ID または名前による値。形式は [付録](Appendix.md#appendix-1-datasheet-field-types) を参照 |

### 戻り値

Array – 更新後の Record

| プロパティ | 型 | 説明 |
| --- | --- | --- |
| id | String | Record ID |
| Fields | Object | 更新後の値。形式は [付録](Appendix.md#appendix-1-datasheet-field-types) を参照 |

### 例

```javascript
const sheet = Application.ActiveSheet
const records = sheet.Record.UpdateRecords({
  Records: [
    {
      id: 'A',
      fields: {
        email: 'demo@qq.com',
        multiSelect: ['1', '2'],
        recordLink: {
          recordIds: ['I', 'K']
        }
      }
    }
  ]
})
```

## CreateRecords()

行を作成します。

### パラメーター

| プロパティ | 型 | 既定値 | 必須 | 説明 |
| --- | --- | --- | --- | --- |
| Records | Array[Object] |  | はい | 新しい Record |

#### Record エントリ

| プロパティ | 型 | 説明 |
| --- | --- | --- |
| id | String | Record ID（該当する場合） |
| Fields | Object | 値。形式は [付録](Appendix.md#appendix-1-datasheet-field-types) を参照 |

### 戻り値

Array – 作成された Record

| プロパティ | 型 | 説明 |
| --- | --- | --- |
| id | String | Record ID |
| Fields | Object | 値。形式は [付録](Appendix.md#appendix-1-datasheet-field-types) を参照 |

### 例

```javascript
const sheet = Application.ActiveSheet
const records = sheet.Record.CreateRecords({
  Records: [
    {
      fields: {
        email: 'demo@qq.com',
        multiSelect: ['1', '2']
      }
    },
    {
      fields: {
        email: 'demo@qq.com',
        multiSelect: ['1', '2']
      }
    }
  ]
})

const records2 = sheet.Record.CreateRecords({
  Records: [
    {
      fields: {
        contact: [
          {
            name: 'yourname',
            nickName: 'yourname',
            id: '88888888',
            avatar_url: 'https://example.com/avatar/user123'
          }
        ]
      }
    }
  ]
})
```

## GetAttachmentURL()

添付ファイルまたは画像のダウンロード URL を取得します。

### パラメーター

注意

`Attachment` を少なくとも渡すか、`UploadId` と `Source` の両方を渡してください。

| プロパティ | 型 | 既定値 | 必須 | 説明 |
| --- | --- | --- | --- | --- |
| Attachment | String |  | いいえ | 添付ファイルのトークン文字列 |
| UploadId | String |  | いいえ | アップロードファイル ID |
| Source | String |  | いいえ | `"upload_ks3"`（ローカルアップロード）または `"cloud"`（クラウドアップロード）である必要あり |

### 戻り値

String – ダウンロード用に開く URL

### 例

```javascript
const sheet = Application.ActiveSheet
const resultURL = sheet.Record.GetAttachmentURL({
  Attachment:
    'IKWRCBAAKA|upload_ks3|image/png|screenshot_20230214165215.png|12070||549*106'
})

const resultURL2 = sheet.Record.GetAttachmentURL({
  UploadId: 'IKWRCBAAKA',
  Source: 'upload_ks3'
})
```
