---
title: "はじめに"
source: https://airsheet.wps.cn/docs/ja-JP/start/quickstart.html
updated: 1789700916000
hash: DYOhDJep
---

# はじめに

## AirScript エディターを開く

1. WPS Docs のホームページで新しいスプレッドシートを作成し、開いて AirScript を体験します。
2. スプレッドシートを開いたら、上部メニューから **効率** - **AirScript エディターツール** をクリックしてエディターページを開きます。
3. 以下の例を順に実行し、結果を確認すると、AirScript をすばやく始められます。

## HelloWorld

### console.log()

```js
console.log("hello world!")
```

### console.error()

```js
console.error("hello world!")
```

## ドキュメント操作

### SheetNames を走査し、対応する Sheet に切り替え、セルを読み取り・変更する

```js
// すべてのワークシート名を走査して出力する
let sheets = Application.Sheets
for (let i = 0; i < sheets.Count; i++) {
    let sheet = sheets.Item(i + 1)
    console.log(sheet.Name) // 各ワークシート名を出力
}

// 現在アクティブな Sheet の名前を出力
console.log(Application.ActiveSheet.Name) // アクティブな Sheet 名を出力

// 現在アクティブなワークシートの A1 セルを取得
let A1 = Application.ActiveSheet.Range("A1")

// A1 セルの内容を出力
console.log(A1.Text) // A1 セルの内容を出力

// A1 セルの内容を変更
A1.Value2 = "bar"
console.log(A1.Text) // 変更後の A1 セルの内容を出力

// A1 セルの背景色を黄色 (255,255,0) に変更
A1.Interior.Color = RGB(255,255,0) // A1 セルを強調表示
```

### ドキュメント内のすべての空でないセルに指定の接尾辞を付加する

```js
// Application.ActiveSheet.UsedRange で使用範囲を取得
const usedRange = Application.ActiveSheet.UsedRange
const startRow = usedRange.Row // 開始行を取得
const startCol = usedRange.Column // 開始列を取得
const endRow = startRow + usedRange.Rows.Count // 終了行を取得
const endCol = startCol + usedRange.Columns.Count // 終了列を取得

for (let i = startRow; i < endRow; i++) {
    const row = Application.ActiveSheet.Rows(i) // 行を特定
    for (let j = startCol; j < endCol; j++) {
        const rg = row.Columns(j) // 行オブジェクトから列を指定してセルを特定
        const text = rg.Text
        // セルが空でなければ foo を付加
        if (text !== '') {
            rg.Value2 = text + 'foo' // セル内容を変更
        }
    }
}

// UsedRange 内の各セルを走査し、内容を出力
for (let i = startRow; i < endRow; i++) {
    for (let j = startCol; j < endCol; j++) {
        const cell = Application.ActiveSheet.Cells(i, j)
        console.log(cell.Text) // すべてのセルデータを出力
    }
}
```

### 新しい Sheet を作成する

```js
// 現在のワークシート名
const currentSheetName = Application.ActiveSheet.Name

// 現在の Sheet の後ろに、既定名のワークシートを 2 枚追加
Application.Sheets.Add(null, Application.Sheets(currentSheetName), 2)

// 現在の Sheet の後ろに「New Worksheet (Right)」という名前の Sheet を追加
const newSheetRight = Application.Sheets.Add(null, Application.Sheets(currentSheetName), 1)
newSheetRight.Name = 'New Worksheet (Right)'

// 現在の Sheet の前に「New Worksheet (Left)」という名前の Sheet を追加
const newSheetLeft = Application.Sheets.Add(Application.Sheets(currentSheetName), null, 1)
newSheetLeft.Name = 'New Worksheet (Left)'
```

## その他

より実践的な例は、[サンプルケース](../example/excel.md) を参照してください。
