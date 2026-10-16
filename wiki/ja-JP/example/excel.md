---
title: "スプレッドシートの例"
source: https://airsheet.wps.cn/docs/ja-JP/example/excel.html
updated: 1789700916000
hash: CM5KTd-p
---

# スプレッドシートの例

AirScript を理解し、スクリプト開発にすぐ取りかかれるよう、スプレッドシートの **スタート** - **クイックツール** にある一部機能のソースコードを掲載しています。 実例のコードを読んだうえで、スクリプトエディターで問題解決に効果的に使えるようになることを期待しています。

## 数式をクリアして値のみを保持

1 つ以上のセルを手動で選択してからスクリプトを実行します。実行後、数式を一括でクリアし、値だけが残ります。

```js
const API = Application

// ユーザー選択
const selection = API.Selection
// ユーザーのアクティブなブックの使用範囲を取得
const usedRange = API.ActiveSheet.UsedRange

// 両方の最小の交差、すなわち走査する最小集合を取得
const rowFrom = Math.max(selection.Row, usedRange.Row)
const rowTo = Math.min(selection.Row + selection.Rows.Count - 1, usedRange.Row + usedRange.Rows.Count - 1)
const colFrom = Math.max(selection.Column, usedRange.Column)
const colTo = Math.min(selection.Column + selection.Columns.Count - 1, usedRange.Column + usedRange.Columns.Count - 1)

for (let i = rowFrom; i <= rowTo; i++) {
  const row = API.ActiveSheet.Rows(i) // 行を決定
  for (let j = colFrom; j <= colTo; j++) {
    const rg = row.Columns(j) // 行オブジェクトから列を指定してセルを決定
    // Text はセルの表示値。元が数式の場合、Text は数式計算後の結果
    // 未計算の結果を取得したい場合は rg.Formula を使用
    const text = rg.Text
    if (text !== '') {
      rg.Value2 = text // 元の Value2 は数式。計算結果で Value2 を書き換えて数式をクリア
    }
  }
}
```

## 無効な電話番号をハイライト

1 つ以上のセルを手動で選択してからスクリプトを実行します。スクリプトはセル内の電話番号が有効かどうかを自動で確認し、実行後に無効な電話番号のセルをハイライトします。

```js
// 電話番号用の正規表現。先頭が 1、2 桁目が 3〜9、残り 9 桁が数字の 11 文字列にマッチ
const PhoneReg = /^1[3-9](\d{9})$/i

// ハイライト色は黄色
const color = RGB(255, 255, 0)

// 簡略化した API 参照
const API = Application

// ユーザーが選択した範囲
const selection = API.Selection

// ブックのアクティブ範囲
const usedRange = API.ActiveSheet.UsedRange

// 両方の最小の交差、すなわち走査する最小集合を取得
const rowFrom = Math.max(selection.Row, usedRange.Row)
const rowTo = Math.min(selection.Row + selection.Rows.Count - 1, usedRange.Row + usedRange.Rows.Count - 1)
const colFrom = Math.max(selection.Column, usedRange.Column)
const colTo = Math.min(selection.Column + selection.Columns.Count - 1, usedRange.Column + usedRange.Columns.Count - 1)

for (let i = rowFrom; i <= rowTo; i++) {
    const row = API.ActiveSheet.Rows(i) // 行オブジェクトを決定
    for (let j = colFrom; j <= colTo; j++) {
        const rg = row.Columns(j) // 続いて列オブジェクト、すなわちセルを決定
        const text = rg.Text  // セルのテキスト内容を取得
        if (text !== '' && !PhoneReg.test(text)) { // 正規表現にマッチしない場合
            rg.Interior.Color = color // セルを黄色でハイライト
        }
    }
}
```

## セルの結合を解除して同じ内容で埋める

1 つ以上のセルを手動で選択してからスクリプトを実行します。スクリプトは結合セルを自動で解除し、同じ内容で埋めます。

```js
const API = Application
// ユーザーが選択した範囲
const selection = API.Selection

// 選択範囲に結合セルがある場合
if (selection.MergeCells) {
  // 結合セル領域をすべて取得
  const areas = selection.MergeArea.Areas
  const areaCount = areas.Count

  // 各結合セル領域を反復処理
  for (let areaIndex = 1; areaIndex <= areaCount; areaIndex++) {
    const area = areas.Item(areaIndex)
    const text = area.Formula // 結合解除前に結合セルの内容を保存

    // 結合セル領域の Range を取得
    const rowFrom = area.Row
    const rowTo = rowFrom + area.Rows.Count - 1
    const colFrom = area.Column
    const colTo = colFrom + area.Columns.Count - 1

    // 結合を解除
    area.UnMerge()
    // 元の結合セル領域に、結合セルの値を書き戻す
    for (let i = rowFrom; i <= rowTo; i++) {
      const row = API.ActiveSheet.Rows(i) // 行を決定
      for (let j = colFrom; j <= colTo; j++) {
        const cell = row.Columns(j) // 列を決定
        cell.Value2 = text // 先に保存した内容をセルに書き込む
      }
    }
  }
}
```

## 重複の出現回数をカウント

1 つ以上のセルを手動で選択してからスクリプトを実行します。スクリプトは選択範囲内の重複する値とその出現回数を自動で集計し、新しいワークシートを作成して、重複する値と件数を新しいワークシートの `A2 セル` と `B2 セル` に書き込みます。

```js
const API = Application
// ユーザーが選択した範囲
const selection = API.Selection
// ブックのアクティブ範囲
const usedRange = API.ActiveSheet.UsedRange

// 両方の最小の交差、すなわち走査する最小集合を取得
const rowFrom = Math.max(selection.Row, usedRange.Row)
const rowTo = Math.min(selection.Row + selection.Rows.Count - 1, usedRange.Row + usedRange.Rows.Count - 1)
const colFrom = Math.max(selection.Column, usedRange.Column)
const colTo = Math.min(selection.Column + selection.Columns.Count - 1, usedRange.Column + usedRange.Columns.Count - 1)

// countMap は各文字列の出現回数を数えるために使用
const countMap = {}
for (let i = rowFrom; i <= rowTo; i++) {
    const row = API.ActiveSheet.Rows(i) // 行を決定
    for (let j = colFrom; j <= colTo; j++) {
        const rg = row.Columns(j) // 続いて列オブジェクト、すなわちセルを決定
        const text = rg.Text // カウント用にセルの値を取得
        if (text) {
            if (countMap[text]) {
                countMap[text]++
            } else {
                countMap[text] = 1
            }
        }
    }
}

const Name = API.ActiveSheet.Name
// 現在のワークシートの後ろに、既定名の新しいワークシートを作成
const newSheet = API.Sheets.Add(null, API.Sheets(Name), 1)

// 新しいワークシートの A1 と B1 に「重複する値」と「出現回数」を書き込む
newSheet.Range('A1').Value2 = '重複する値'
newSheet.Range('B1').Value2 = '出現回数'

// 2 行目から、1 列目に文字列、2 列目に出現回数を出力
let index = 2
for (let k in countMap) {
    const row = newSheet.Rows(index)
    row.Columns(1).Value2 = "'" + k
    row.Columns(2).Value2 = countMap[k] + ''
    index++;
}
```
