---
title: "付録"
source: https://airsheet.wps.cn/docs/ja-JP/api/excel/databook/Appendix.html
updated: 1789700916000
hash: CD6n301F
---

# 付録

## 付録 1: データシートの Field の種類

| フィールドの種類 | 型 | Field 作成時の形状 | 書き込み値 | 読み取り値 |
|---|---|---|---|---|
| 複数行テキスト | MultiLineText | 追加ルールなし | String | String |
| 日付 | Date | 追加ルールなし | String、`yyyy/mm/dd` 形式 | String |
| 時刻 | Time | 追加ルールなし | String、`hh:mm:ss` 形式 | String |
| 数値 | Number | 追加ルールなし | Number | Number |
| 通貨 | Currency | 追加ルールなし | Number | Number |
| パーセント | Percentage | 追加ルールなし | Number | Number |
| 身分証明書番号 | ID | 追加ルールなし | String（ID ルールに一致する必要あり） | String |
| 電話番号 | Phone | 追加ルールなし | String（電話番号ルールに一致する必要あり） | String |
| メール | Email | 追加ルールなし | String（メールルールに一致する必要あり） | String |
| URL | Url | 任意の `displayText`：`{"name":"Link","type":"Url","displayText":"Open"}` | String（有効な URL） | String |
| チェックボックス | Checkbox | 追加ルールなし | true / false | Boolean |
| 単一選択 | SingleSelect | オプションを 1 つ以上：`{"name": "Option","type": "SingleSelect","items": [{ "value": "item1" }]}` | String（いずれかのオプションと一致する必要あり） | String |
| 複数選択 | MultipleSelect | オプションを 1 つ以上：`{"name": "Option","type": "SingleSelect","items": [{ "value": "item1" }, { "value": "item2" }]}` | String 配列 | String 配列 |
| 評価 | Rating | `max` は 1〜5：`{"name": "Rating","type": "Rating","max": 5}` | Number（0 より大きく max 未満） | Number |
| 進捗 | Complete | 追加ルールなし | Number（0〜100） | String |
| 連絡先 | Contact | `multipleContacts`（`<bool>`）、`noticeNewContact`（`<bool>`）：`{"name": "Contact","type": "Contact","multipleContacts": false,"noticeNewContact": false}` | 書き込み不可 | Object |
| 添付ファイル | Attachment | 追加ルールなし | 書き込み不可 | Object |
| リンク | Link | `linkSheet`、`multipleLinks`：`{"name": "Link","type": "Link","linkSheet": "<sheetId>","multipleLinks": false}` | リンク先 Sheet の行の配列 |  |
| リッチテキスト | Note | 追加ルールなし | 書き込み不可 | Object |
| 自動採番 | AutoNumber | 追加ルールなし | 書き込み不可 | Number |
| 作成者 | CreatedBy | 追加ルールなし | 書き込み不可 | Object |
| 作成日時 | CreatedTime | 追加ルールなし | 書き込み不可 | String |
| 数式 | Formula | 追加ルールなし | 書き込み不可 | 数式に依存 |
| 参照 | Lookup | 追加ルールなし | 書き込み不可 | 参照先 Field と同じ形状 |

## 付録 2: データシートの View の種類

| View の種類 | 説明 |
|---|---|
| Grid | グリッド View |
| Kanban | カンバン View |
| Gallery | ギャラリー View |
| Form | フォーム View |
| Gantt | ガント View |

## 付録 3: フィルター条件

フィルターは行を絞り込みます。構成は 2 部です。`mode` は条件の結合方法、`criteria` は各ルール（`field`、`op`、`values`）を保持します。

```json
{
  "mode": "AND",
  "criteria": [
    {
      "field": "Name",
      "op": "Intersected",
      "values": ["Datasheet", "12345"]
    },
    {
      "field": "Quantity",
      "op": "Greater",
      "values": ["1"]
    }
  ]
}
```

| 演算子 | 意味 |
|---|---|
| Equals | 等しい |
| NotEqu | 等しくない |
| Greater | より大きい |
| GreaterEqu | 以上 |
| Less | より小さい |
| LessEqu | 以下 |
| GreaterEquAndLessEqu | 範囲内（両端を含む） |
| LessOrGreater | 範囲外（両端を含まない） |
| BeginWith | 前方一致 |
| EndWith | 後方一致 |
| Contains | 含む |
| NotContains | 含まない |
| Intersected | 列挙値のいずれか |
| Empty | 空である |
| NotEmpty | 空でない |

各演算子は、`values` に含められる要素数の上限を定めます。上限を超えると、そのルールは無視されます。

`Empty` と `NotEmpty` は値を持てません。「範囲」系の演算子は最大 2 個まで、`Intersected` は最大 65535 個まで、その他は最大 1 個までです。

注意

`filter` は構造体ではありません。`field` が欠落している、`op` / `values` が無効、`values` が長すぎる、などルールが無効な場合は、リクエスト全体が失敗します。

動的日付フィルターでは、`values` にオブジェクト要素を使用します。

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

動的日付では、`values` 内の各オブジェクトは `type` を `DynamicSimple` に設定する必要があります。`op` が `Equals` のとき、`dynamicType` は次の値にできます（大文字・小文字を区別しません）。

| 値 | 意味 |
|---|---|
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

`op` が `greater` または `less` のとき、`dynamicType` は昨日・今日・明日のみにできます。
