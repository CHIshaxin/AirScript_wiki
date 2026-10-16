---
title: "ベストプラクティス"
source: https://airsheet.wps.cn/docs/ja-JP/start/best-practice.html
updated: 1789700916000
hash: TgR3BlWU
---

# ベストプラクティス

## ユーザー体験の最適化

#### ハイパーリンクでスクリプトを実行する

スクリプトはハイパーリンクから起動できます。バインド方法は次のとおりです。

1. **挿入** - **リンク** を開き、リンク設定ページを表示します。
2. テキスト入力欄に、スクリプトの機能を表す適切な表示文言を入力します。
3. **種類** の横のドロップダウンをクリックし、**AirScript スクリプト** を選び、バインドするスクリプトを指定します。
4. **OK** をクリックします。

挿入したハイパーリンクをクリックするとスクリプトが実行されます。

この方法は WYSIWYG に近い体験になり、スクリプトエディターを開いて実行ボタンを押すより、はるかに分かりやすくなります。

#### より直感的な実行体験

スクリプトで特定のセル範囲を扱う場合、実装方法は複数あります。`Application.Selection` を通じてユーザーが選択したセルを読み取り・操作するのが、最も直感的なユーザー体験です。

以下は、スクリプト内の変数で走査範囲を定義する例です。範囲が変わるたびに 4 つの変数を書き換える必要があり、不便で体験も悪くなります。

```javascript
// 悪い例
const startRow = 1
const startColumn = 1
const endRow = 3
const endColumn = 3
for(let i = startRow; i <= endRow; i++){
  let row=Application.Rows(i)
  for (let j = startColumn; j <= endColumn; j++){
    console.log(row.Columns(j).Text)
  }
}
```

以下は `Application.Selection` で、ドキュメント上でユーザーが選択したセル範囲を取得する例です。手軽で直感的です。

```javascript
// 良い例
let selection = Application.Selection
let startRow = selection.Row
let startCol = selection.Column
let endRow = startRow + selection.Rows.Count - 1
let endCol = startCol + selection.Columns.Count - 1

for (let i = startRow; i <= endRow; i++) {
    let row = Application.Rows(i)
    for (let j = startCol; j <= endCol; j++) {
        console.log(row.Columns(j).Text)
    }
}
```

## パフォーマンスの最適化

#### オブジェクトの再利用を最大限に行う

`Application` への各関数呼び出しは、スクリプトエンジンがファイルデータを操作するため、繰り返し呼び出すとパフォーマンスが無駄になります。

たとえば 100×100 セルの内容を走査して出力する場合、オブジェクトを再利用すると処理速度がおよそ倍になります。

オブジェクトを再利用しない場合：

```javascript
// 悪い例
let start = new Date()
for(let i= 1;i<=100;i++){
  for (let j =1;j<=100;j++){
    Application.Rows(i).Columns(j).Text
  }
}
console.log(new Date()-start,'ms')
```

より良い方法：

```javascript
// 良い例
let start = new Date()
for(let i= 1;i<=100;i++){
  let row=Application.Rows(i)
  for (let j =1;j<=100;j++){
    row.Columns(j).Text
  }
}
console.log(new Date()-start)
```

#### UsedRange で走査範囲を絞り込む

上記の [より直感的な実行体験](#より直感的な実行体験) の例では、 スプレッドシート全体を選択すると反復は兆単位に達し、行全体を 1 行だけ選んだ場合でも数百万回の反復になります。 しかし実際に使っている範囲はそれほど大きくありません。`Application.ActiveSheet.UsedRange` でワークブックの使用範囲を把握すれば、スクリプトの実行時間を大きく短縮できます。

以下は、その例を改善したコードです。

```javascript
const selection = Application.Selection
const usedRange = Application.ActiveSheet.UsedRange

// 走査範囲を決定
const rowFrom = Math.max(selection.Row, usedRange.Row)
const rowTo = Math.min(selection.Row + selection.Rows.Count - 1, usedRange.Row + usedRange.Rows.Count - 1)
const colFrom = Math.max(selection.Column, usedRange.Column)
const colTo = Math.min(selection.Column + selection.Columns.Count - 1, usedRange.Column + usedRange.Columns.Count - 1)

for (let i = rowFrom; i <= rowTo; i++) {
    const row = Application.ActiveSheet.Rows(i) // オブジェクトを再利用
    for (let j = colFrom; j <= colTo; j++) {
        console.log(row.Columns(j).Text)
    }
}
```
