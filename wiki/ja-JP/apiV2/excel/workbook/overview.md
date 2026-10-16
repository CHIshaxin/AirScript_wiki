---
title: "ワークブック API"
source: https://airsheet.wps.cn/docs/ja-JP/apiV2/excel/workbook/overview.html
updated: 1789700916000
hash: -g51xr4w
---

# ワークブック API

ワークブック API は、セル操作、書式設定、グラフ作成、データフィルタリングなど、Excel ワークシートを完全に制御する機能を提供します。

## コアオブジェクト

- **Application** - グローバルメソッドとプロパティを提供するアプリケーションオブジェクト
- **Workbook** - ワークシートコレクションを管理するワークブックオブジェクト
- **Worksheet** - 個々のシートを操作するワークシートオブジェクト
- **Range** - セル範囲を操作する Range オブジェクト
- **Chart** - グラフを作成・操作するチャートオブジェクト

## クイックスタート

```javascript
// 現在のアクティブシートを取得
const sheet = ActiveSheet;

// セル値を読み取る
const value = sheet.Cells(1, 1).Value2;

// セル値を設定する
sheet.Cells(1, 1).Value2 = "Hello World";

// 範囲を取得する
const range = sheet.Range("A1:B10");
```

詳細なドキュメントについては、サイドバーナビゲーションから特定の API オブジェクトを選択してください。
