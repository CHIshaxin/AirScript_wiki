---
title: "ユースケース"
source: https://airsheet.wps.cn/docs/ja-JP/apitoken/scenario.html
updated: 1789700916000
hash: B6APYY3x
---

# ユースケース

スクリプトトークンのおかげで、スマートスプレッドシートの強力な機能が十分に活かされ、開発者はより効率的かつ正確にタスクを実行できます。Web スクレイピング、データ分析、自動化プロセスのいずれでも、スクリプトトークンを使って実現できます。

以下は、実際のシーンをもとに作成した例と説明です。開発者の皆さまの参考になれば幸いです。

## 1. プライベート情報クエリ

現代社会では個人のプライバシーがますます重視されています。その傾向はさまざまな面に表れています。

第一に、教育の分野では、多くの学校が在籍中の学生情報の保護を強化しており、学生の個人情報を第三者に販売・共有することを禁じています。これには学生の**家族情報、学歴、試験の成績**などが含まれます。

第二に、医療の分野では、患者のプライバシー保護が重要な課題となっています。医師その他の医療従事者は厳格なプライバシー規制を遵守し、**患者の個人情報**が漏えいしないようにする必要があります。

さらに、ソーシャルメディアの分野でも、多くのプラットフォームがユーザー情報の保護を強化し始めています。より厳格なデータセキュリティ対策を講じ、ユーザーデータが漏えいしたり悪用されたりしないようにしています。

以下では、学生の成績照会のオンライン例を示します。スクリプトトークンを用いて実装したプライベート情報照会機能を、オンラインで体験できます。既存の学生成績シートは次のとおりです。

上記のオンライン例の実装では、**スクリプトトークンで AirScript スクリプトを実行**し、スプレッドシート内のデータを照会してインターフェースに返して表示しています。スクリプトの例は次のとおりです。

```js
// 学生IDを取得
const student_id = '1'
// 学生名を取得
const student_name = 'ジン・シャオメイ'
// 使用範囲の行数を取得
const rowCount = Application.ActiveSheet.UsedRange.Rows.Count

// 一致する学生データを取得
const data = []
for (let i = 1; i <= rowCount; i++) {
  const studentId = Application.ActiveSheet.Range(`A${i}`).Value2
  const studentName = Application.ActiveSheet.Range(`B${i}`).Value2

  if (studentId === student_id && studentName === student_name) {
    const sex = Application.ActiveSheet.Range(`C${i}`).Value2
    const className = Application.ActiveSheet.Range(`D${i}`).Value2
    const language = Application.ActiveSheet.Range(`E${i}`).Value2
    const math = Application.ActiveSheet.Range(`F${i}`).Value2
    const english = Application.ActiveSheet.Range(`G${i}`).Value2
    const literature = Application.ActiveSheet.Range(`H${i}`).Value2
    const total = Application.ActiveSheet.Range(`I${i}`).Value2
    const rate = Application.ActiveSheet.Range(`J${i}`).Value2

    data.push({
      id: studentId,
      name: studentName,
      sex: sex,
      className: className,
      language: language,
      math: math,
      english: english,
      literature: literature,
      total: total,
      rate: rate
    })
    break
  }
}

// 一致する学生データを返す
return data
```

## 2. ECデータ同期

EC データの同期は、EC プラットフォームが販売・運用の両面で効率的に動くために重要な側面です。データソースとしては、データベース、ERP、CRM など複数のシステムにまたがる場合があります。開発者がまず行うのは、商品情報・注文情報・物流情報などを手動でエクスポートするか、webhook 経由で取得することです。次に自社サーバー上でそのデータを受け取り、クレンジングと変換を行い、最後にスクリプトトークンを通じて WPS Docs のスマートスプレッドシートに書き込み、データ同期を完了します。
