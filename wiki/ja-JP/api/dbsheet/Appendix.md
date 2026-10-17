---
title: "付録"
source: https://airsheet.wps.cn/docs/ja-JP/api/dbsheet/Appendix.html
updated: 1789700916000
hash: zB0K1UB8
---

# 付録

## 付録 1: 多次元テーブルのフィールド型

[テーブルを迅速に作成](../../example/dbSheet.md#quickly-create-a-table) も参照してください。

| フィールド型 | 型 | フィールド作成時の形状 | 書き込み値 | 読み取り値 |
| --- | --- | --- | --- | --- |
| 複数行テキスト | MultiLineText | 追加ルールなし | 文字列 | 文字列 |
| 日付 | Date | 追加ルールなし | 文字列、`yyyy/mm/dd` 形式 | 文字列 |
| 時刻 | Time | 追加ルールなし | 文字列、`hh:mm:ss` 形式 | 文字列 |
| 数値 | Number | 追加ルールなし | 数値 | 数値 |
| 通貨 | Currency | 追加ルールなし | 数値 | 数値 |
| パーセンテージ | Percentage | 追加ルールなし | 数値 | 数値 |
| 身分証明書 | ID | 追加ルールなし | 文字列（ID ルールに一致する必要あり） | 文字列 |
| 電話番号 | Phone | 追加ルールなし | 文字列（電話番号ルールに一致する必要あり） | 文字列 |
| メール | Email | 追加ルールなし | 文字列（メール形式ルールに一致する必要あり） | 文字列 |
| URL | Url | 任意の `displayText`: `{"name":"Link","type":"Url","displayText":"Open"}` | 文字列（有効な URL） | 文字列 |
| チェックボックス | Checkbox | 追加ルールなし | true / false | Boolean |
| 単一選択 | SingleSelect | 少なくとも 1 つのオプション: `{"name": "Option","type": "SingleSelect","items": [{ "value": "item1" }]}` | 文字列（いずれかのオプションと一致する必要あり） | 文字列 |
| 複数選択 | MultipleSelect | 少なくとも 1 つのオプション: `{"name": "Option","type": "SingleSelect","items": [{ "value": "item1" }, { "value": "item2" }]}` | 文字列の配列 | 文字列の配列 |
| 評価 | Rating | `max` は 1〜5: `{"name": "Rating","type": "Rating","max": 5}` | 数値、0 より大きく max 未満 | 数値 |
| 進捗 | Complete | 追加ルールなし | 数値、0〜100 | 文字列 |
| 連絡先 | Contact | `multipleContacts`（`<bool>`）、`noticeNewContact`（`<bool>`）: `{"name": "Contact","type": "Contact","multipleContacts": false,"noticeNewContact": false}` | 書き込み不可 | オブジェクト |
| 添付ファイル | Attachment | 追加ルールなし | 書き込み不可 | オブジェクト |
| リンク | Link | `linkSheet`、`multipleLinks`: `{"name": "Link","type": "Link","linkSheet": "<sheetId>","multipleLinks": false}` | リンク先 Sheet の行の配列 |  |
| リッチテキスト | Note | 追加ルールなし | 書き込み不可 | オブジェクト |
| 自動採番 | AutoNumber | 追加ルールなし | 書き込み不可 | 数値 |
| 作成者 | CreatedBy | 追加ルールなし | 書き込み不可 | オブジェクト |
| 作成日時 | CreatedTime | 追加ルールなし | 書き込み不可 | 文字列 |
| 数式 | Formula | 追加ルールなし | 書き込み不可 | 数式に依存 |
| ルックアップ | Lookup | 追加ルールなし | 書き込み不可 | 参照フィールドと同じ形状 |

## 付録 2: 多次元テーブルの View 型

[テーブルを迅速に作成](../../example/dbSheet.md#quickly-create-a-table) も参照してください。

| View 型 | 説明 |
| --- | --- |
| Grid | グリッド View |
| Kanban | カンバン View |
| Gallery | ギャラリー View |
| Form | フォーム View |
| Gantt | ガント View |

## 付録 3: フィルター条件

フィルターは行を制限します。2 つの部分があります。`mode` は条件の結合方法、`criteria` が各ルール（`field`、`op`、`values`）を保持します。

```json
{
  "mode": "AND",
  "criteria": [
    {
      "field": "Name",
      "op": "Intersected",
      "values": ["Multidimensional Table", "12345"]
    },
    {
      "field": "Quantity",
      "op": "greater",
      "values": ["1"]
    }
  ]
}
```

| 演算子 | 意味 |
| --- | --- |
| Equals | 等しい |
| NotEqu | 等しくない |
| Greater | より大きい |
| GreaterEqu | 以上 |
| Less | より小さい |
| LessEqu | 以下 |
| GreaterEquAndLessEqu | 範囲内（両端を含む） |
| LessOrGreater | 範囲内（両端を含まない） |
| BeginWith | 前方一致 |
| EndWith | 後方一致 |
| Contains | 含む |
| NotContains | 含まない |
| Intersected | 列挙値のいずれか |
| Empty | 空である |
| NotEmpty | 空でない |

各演算子は、`values` に含められる要素数の上限を定めます。上限を超えると、そのルールは無視されます。

`Empty` と `NotEmpty` は値を持てません。「範囲」系の演算子は最大 2 個まで、`Intersected` は最大 65535 個、その他は最大 1 個です。

備考

`filter` は構造体ではありません。`field` が欠落している、`op` / `values` が無効、`values` が長すぎる、などルールが無効な場合は、リクエスト全体が失敗します。

動的日付フィルターでは、`values` の要素にオブジェクトを使います。

```json
{
  "mode": "AND",
  "criteria": [
    {
      "field": "Date",
      "op": "Equals",
      "values": [
        {
          "dynamicType": "lastMonth",
          "type": "DynamicSimple"
        }
      ]
    }
  ]
}
```

ヒント

上記の例は「先月と等しい」という意味です。

動的日付では、`values` 内の各オブジェクトは `type` を `DynamicSimple` に設定する必要があります。`op` が `Equals` のとき、`dynamicType` は次の値にできます（大文字・小文字は区別しません）。

| 値 | 意味 |
| --- | --- |
| today | 今日 |
| yesterday | 昨日 |
| tomorrow | 明日 |
| last7Days | 過去 7 日間 |
| last30Days | 過去 30 日間 |
| thisWeek | 今週 |
| lastWeek | 先週 |
| nextWeek | 来週 |
| thisMonth | 今月 |
| lastMonth | 先月 |
| nextMonth | 来月 |

ヒント

`op` が `greater` または `less` のとき、`dynamicType` に使えるのは yesterday、today、tomorrow のみです。
