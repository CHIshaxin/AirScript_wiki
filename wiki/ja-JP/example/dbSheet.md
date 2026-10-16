---
title: "マルチディメンショナルテーブルの例"
source: https://airsheet.wps.cn/docs/ja-JP/example/dbSheet.html
updated: 1789700916000
hash: DbK_OZya
---

# マルチディメンショナルテーブルの例

マルチディメンショナルテーブル開発向けのサンプルスクリプトです。素早く理解し、スクリプト開発に取りかかれるよう役立ててください。

## 選択範囲に値を一括入力

```javascript
function main() {
  var activeView = Application.Selection.GetActiveView()

  var selectedRecords = Application.Selection.GetSelectionRecords()[0]
  var time = getNowTime()
  var date = getNowDate()

  // 備考: ここの「日付」「時刻」「カテゴリ」は、実際のテーブルのフィールド名に置き換えてください
  Application.Record.UpdateRecords({
    SheetId: activeView.sheetId,
    Records: selectedRecords.map(item => ({
      id: item.id,
      fields: {
        日付: date,
        時刻: time,
        カテゴリ: 'B'
      }
    }))
  })
}
// 現在時刻を "hh:mm:ss" 形式で取得
function getNowTime() {
  return new Date().toTimeString().split(' ')[0]
}
// 現在日付を "yyyy:MM:dd" 形式で取得
function getNowDate() {
  var date = new Date()
  return date.getFullYear() + '/' + (date.getMonth() + 1) + '/' + date.getDate()
}

main()
```

## 「ワンクリックアーカイブ」のクイック実装

次のコードは、1 つのファイル内に同じデータ構造を持つ 2 つのテーブルを用意し、テーブル 1 から完了したデータをテーブル 2 に挿入し、テーブル 1 から該当データを削除します。

```javascript
function main() {
  var sheets = Application.Sheet.GetSheets()
  // チェック済みかつカテゴリ B のデータをフィルタ
  var finishedRecords = Application.Record.GetRecords({
    SheetId: sheets[0].id,
    Filter: {
      mode: 'AND',
      criteria: [
        {
          field: '完了',
          op: 'Equals',
          values: ['1']
        },
        {
          field: 'カテゴリ',
          op: 'Equals',
          values: ['B']
        }
      ]
    }
  })
  // フィルタ結果がある場合
  if (finishedRecords) {
    // テーブル 2 にデータを挿入
    Application.Record.CreateRecords({
      SheetId: sheets[1].id,
      Records: finishedRecords.records.map(item => ({
        fields: item.fields
      }))
    })
    // テーブル 1 からデータを削除
    Application.Record.DeleteRecords({
      SheetId: sheets[0].id,
      RecordIds: finishedRecords.records.map(item => item.id)
    })
  }
}
main()
```

ヒント

上記 2 つの例を組み合わせると、アーカイブ日時の自動設定や、選択レコードのワンクリックアーカイブなどの機能を実装できます。

## テーブルを素早く作成

```javascript
function main() {
  Application.Sheet.CreateSheet({
    Name: 'マイテーブル',
    Fields: [
      { name: '名前', type: 'MultiLineText' },
      { name: '数量', type: 'Number' },
      { name: '日付', type: 'Date' },
      { name: '時刻', type: 'Time' },
      { name: 'チェックボックス', type: 'Checkbox' },
      { name: 'ハイパーリンク', type: 'Url' },
      { name: '評価', type: 'Rating', max: 5 },
      { name: '電話', type: 'Phone' },
      { name: '身分証', type: 'ID' },
      { name: '通貨', type: 'Currency' },
      { name: 'パーセンテージ', type: 'Percentage' },
      { name: 'メール', type: 'Email' },
      { name: '進捗', type: 'Complete' },
      {
        name: 'カテゴリ',
        type: 'SingleSelect',
        items: [{ value: 'A' }, { value: 'B' }, { value: 'C' }]
      },
      {
        name: 'ステータス',
        type: 'MultipleSelect',
        items: [{ value: '完了' }, { value: '未着手' }, { value: '進行中' }]
      },
      {
        name: '連絡先',
        type: 'Contact',
        multipleContacts: false,
        noticeNewContact: false
      },
      { name: 'リッチテキスト', type: 'Note' },
      { name: '添付', type: 'Attachment' },
      { name: '数式', type: 'Formula' },
      { name: '作成日時', type: 'CreatedTime' },
      { name: '作成者', type: 'CreatedBy' },
      { name: '自動採番', type: 'AutoNumber' }
    ],
    Views: [
      { name: 'Grid', type: 'Grid' },
      { name: 'Kanban', type: 'Kanban' },
      { name: 'Gallery', type: 'Gallery' },
      { name: 'Form', type: 'Form' },
      { name: 'Gantt', type: 'Gantt' }
    ]
  })
}
main()
```

## フォーマットされたデータの一括挿入

```javascript
function main() {
  var template = ['Product', 10]
  var records = []
  // カテゴリ A の商品を 100 件挿入
  for (let i = 1; i < 100; i++) {
    records.push({
      名前: template[0] + i,
      数量: template[1],
      カテゴリ: 'A'
    })
  }
  for (let i = 1; i < 100; i++) {
    records.push({
      名前: template[0] + (100 + i),
      数量: template[1] + 10,
      カテゴリ: 'B'
    })
  }
  for (let i = 1; i < 100; i++) {
    records.push({
      名前: template[0] + (200 + i),
      数量: template[1] + 10,
      カテゴリ: 'C'
    })
  }
  var sheet = Application.Selection.GetActiveSheet()
  Application.Record.CreateRecords({
    SheetId: sheet.sheetId,
    Records: records.map(item => ({
      fields: item
    }))
  })
}
main()
```
