---
title: "メール API"
source: https://airsheet.wps.cn/docs/ja-JP/api/advanced/SMTP.html
updated: 1789700916000
hash: YgpCcvbj
---

# メール API

外部メールサービス経由でメールを送信します。

### クイックスタート

```javascript
// ログイン
let mailer = SMTP.login({
    host: "smtp.example.com", // ホスト名
    port: 465, // ポート
    secure: true, // TLS
    username: "sender@example.com", // アカウント
    password: "Pa55W0rd" // パスワード
})
// メール送信
mailer.send({
    from: "sender@example.com", // 差出人
    to: "reciever@example.com", // 宛先
    subject: "this is subject.", // 件名
    text: "this is text.", // プレーンテキスト
    html: `<p> this is html </p>` // HTML
})
// 表示名を指定可能
mailer.send({
    from: "Administrator <admin@example.com>",
    to: "Recipient <username@example.com>",
    subject: "this is subject.",
    text: "this is text.",
    html: `<p> this is html </p>`
})
// 複数宛先
mailer.send({
    from: "Administrator <admin@example.com>",
    to: ["username1@example.com","Recipient 2 <username2@example.com>"],
    subject: "this is subject.",
    text: "this is text.",
    html: `<p> this is html </p>`
})
```

## SMTP

### メソッド

| メソッド | 戻り値の型 | 概要 |
|---|---|---|
| [login(argvs)](#login-argvs) | [Mailer](#mailer) | ログインし、メール送信オブジェクトを返す |

## login(argvs)

ログインし、メール送信オブジェクトを返します。

```javascript
// Gmail でログイン
let mailer = SMTP.login({
    host: "smtp.gmail.com", // Gmail の SMTP ホスト
    port: 465,
    username: "1000000000@gmail.com", // Gmail アドレス
    password: "xxxxxxxxxxxx", // SMTP パスワード（Gmail アカウントのパスワードではない）
    secure: true
});
```

### パラメーター

| パラメーター | 型 | 既定値 | 必須 | 説明 |
|---|---|---|---|---|
| argvs | LoginArgvs | undefined | はい | SMTP 設定を含む JavaScript オブジェクト。後述 |

### LoginArgvs

| パラメーター | 型 | 既定値 | 必須 | 説明 |
|---|---|---|---|---|
| host | string | undefined | はい | メールサーバーのホスト名 |
| port | number | undefined | はい | SMTP ポート。`host` が undefined の場合、`secure` に応じた既定値: false なら 587、true なら 465。 |
| secure | boolean | undefined | はい | TLS を使うか。ポート 465 の場合は通常 true。587 または 25 の場合は false。 |
| username | string | undefined | はい | 認証用アカウント名 |
| password | string | undefined | はい | 認証用パスワード |
| timeout | number | 10000 | いいえ | 接続タイムアウト（ミリ秒） |

### 戻り値

[Mailer](#mailer) - メール送信オブジェクト

## Mailer

[login(argvs)](#login-argvs) で作成され、メール送信に使います。

### メソッド

| メソッド | 戻り値の型 | 概要 |
|---|---|---|
| [send(message)](#send-message) | undefined | メールを送信する |

## send(message)

メールを送信します。

```javascript
mailer.send({
    from: ["Administrator <admin@example.com>"],
    to: ["username@example.com", "UserName <username2@example.com>"],
    subject: "this is subject.",
    text: "this is text.",
    html: `<p> this is html </p>`
})
```

### パラメーター

| パラメーター | 型 | 既定値 | 必須 | 説明 |
|---|---|---|---|---|
| message | messageArgvs | undefined | はい | メール内容を含む JavaScript オブジェクト。後述 |

### messageArgvs

| パラメーター | 型 | 既定値 | 必須 | 説明 |
|---|---|---|---|---|
| from | string | undefined | はい | 差出人メールアドレス |
| to | string / string[] | undefined | はい | 宛先メールアドレス（複数可） |
| subject | string | undefined | はい | 件名 |
| text | string | undefined | はい | プレーンテキスト本文 |
| html | string | undefined | いいえ | HTML 本文 |
