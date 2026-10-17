---
title: "ネットワーク API"
source: https://airsheet.wps.cn/docs/ja-JP/api/advanced/HTTP.html
updated: 1789700916000
hash: CqxPYc2e
---

# ネットワーク API

AirScript はグローバルな HTTP オブジェクトを公開します。開発者はそのメソッドを呼び出して外部サービスにリクエストし、成功時にはサーバーのレスポンスが同期的に返されます。

この API はブラウザの `fetch()` に近い挙動のため、フロントエンド開発者でもすぐに使い始められます。

ヒント

HTTP オブジェクトでリクエストを送る前に、スクリプトエディターのツールバー **サービス** メニューから `Network API` サービスを追加してください。

### クイックスタート

```javascript
// ネットワークリクエストを送信する
const resp = HTTP.fetch('https://jsonplaceholder.typicode.com/posts/1', {
  timeout: 2000
})
const data = resp.json()
console.log(data.note, data.content)
```

### メソッド

| メソッド | 戻り値の型 | 概要 |
| --- | --- | --- |
| [fetch(url[, options])](#fetch) | [Response](#response) | 任意の HTTP リクエストを送信する |
| [get(url[, options])](#get) | [Response](#response) | GET リクエストを送信する |
| [delete(url[, options])](#delete) | [Response](#response) | DELETE リクエストを送信する |
| [post(url,body[, options])](#post) | [Response](#response) | POST リクエストを送信する |
| [put(url,body[, options])](#put) | [Response](#response) | PUT リクエストを送信する |

## fetch(url[, options])

HTTP リクエストを送信します。ヘッダーやボディを設定できます。

### 例

```javascript
const resp = HTTP.fetch('https://www.wps.com', {
  method: 'GET',
  timeout: 2000,
  headers: {
    'User-Agent':
      'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/109.0.0.0 Safari/537.36'
  }
})
console.log(resp.text())
```

### パラメーター

| パラメーター | 型 | 既定値 | 必須 | 説明 |
| --- | --- | --- | --- | --- |
| url | string |  | はい | リクエスト先の URL。ポート番号を含まないホスト名のみ使用可能 |
| options | [RequestOption](#requestoption) | undefined | いいえ | 任意のリクエスト設定オブジェクト。後述。 |

### RequestOption

| パラメーター | 型 | 既定値 | 必須 | 説明 |
| --- | --- | --- | --- | --- |
| [method](#method) | string | GET | いいえ | HTTP メソッド。例: `GET`, `POST`, `PUT`, `DELETE` など。 |
| [timeout](#timeout) | number | 10000 | いいえ | タイムアウト（ミリ秒）。範囲 0〜60000。範囲外は 10 秒にフォールバック。 |
| [headers](#headers) | object | undefined | いいえ | リクエストヘッダー。例: `cookie`。 |
| [body](#body) | string | undefined | いいえ | リクエストボディ。 |

### 戻り値

[Response](#response) - サーバーからのレスポンス

## get(url[, options])

GET リクエストを送信します。

### 例

```javascript
const resp = HTTP.get('https://reqres.in/api/users/2')
console.log(resp.json())
```

### パラメーター

| パラメーター | 型 | 既定値 | 必須 | 説明 |
| --- | --- | --- | --- | --- |
| url | string |  | はい | リクエスト先の URL。ポート番号を含まないホスト名のみ使用可能 |
| options | [MethodRequestOption](#methodrequestoption) | undefined | いいえ | このリクエスト用の任意の設定オブジェクト。後述。 |

### MethodRequestOption

| パラメーター | 型 | 既定値 | 必須 | 説明 |
| --- | --- | --- | --- | --- |
| [timeout](#timeout) | number | 10000 | いいえ | タイムアウト（ミリ秒）。範囲 0〜60000。範囲外は 10 秒にフォールバック。 |
| [headers](#headers) | object | undefined | いいえ | リクエストヘッダー。例: `cookie`。 |

### 戻り値

[Response](#response) - サーバーからのレスポンス

## delete(url[, options])

DELETE リクエストを送信します。

### 例

```javascript
const resp = HTTP.delete('https://reqres.in/api/users/2')
console.log(resp.status)
```

### パラメーター

| パラメーター | 型 | 既定値 | 必須 | 説明 |
| --- | --- | --- | --- | --- |
| url | string |  | はい | リクエスト先の URL。ポート番号を含まないホスト名のみ使用可能 |
| options | [MethodRequestOption](#methodrequestoption) | undefined | いいえ | このリクエスト用の任意の設定オブジェクト。後述。 |

### 戻り値

[Response](#response) - サーバーからのレスポンス

## post(url,body[, options])

POST リクエストを送信します。

### 例

```javascript
// フォームを送信する
const formResp = HTTP.post(
  'https://www.example.com',
  { foo: 'bar' },
  { headers: { 'content-type': 'multipart/form-data' } }
)

// JSON を送信する
const resp = HTTP.post('https://reqres.in/api/users', {
  name: 'morpheus',
  job: 'leader'
})

console.log(resp.json())
```

### パラメーター

| パラメーター | 型 | 既定値 | 必須 | 説明 |
| --- | --- | --- | --- | --- |
| url | string |  | はい | リクエスト先の URL。ポート番号を含まないホスト名のみ使用可能 |
| body | string\| object |  | はい | リクエストボディ |
| options | [MethodRequestOption](#methodrequestoption) | undefined | いいえ | このリクエスト用の任意の設定オブジェクト。後述。 |

### 戻り値

[Response](#response) - サーバーからのレスポンス

## put(url,body[, options])

PUT リクエストを送信します。

### 例

```javascript
const resp = HTTP.put('https://reqres.in/api/users/200', {
  name: 'wps',
  job: 'developer'
})
console.log(resp.json())
```

### パラメーター

| パラメーター | 型 | 既定値 | 必須 | 説明 |
| --- | --- | --- | --- | --- |
| url | string |  | はい | リクエスト先の URL。ポート番号を含まないホスト名のみ使用可能 |
| body | string\| object |  | はい | リクエストボディ |
| options | [MethodRequestOption](#methodrequestoption) | undefined | いいえ | このリクエスト用の任意の設定オブジェクト。後述。 |

### 戻り値

[Response](#response) - サーバーからのレスポンス

## Response

HTTP リクエスト後の戻り値。レスポンスはストリームに近く、`text()`、`json()`、`binary()` のうち最初に呼んだ 1 回だけがデータを返します。

### 例

```javascript
let resp = HTTP.get('https://jsonplaceholder.typicode.com/posts/1')
console.log(resp.status) // 200
console.log(resp.statusText) // OK
console.log(resp.text()) // `{foo:"bar"}`
console.log(resp.json()) // {foo:"bar"}
console.log(resp.status) // [...]
```

### メソッド

| メソッド | 戻り値の型 | 概要 |
| --- | --- | --- |
| [status](#status) | number | HTTP ステータスコード |
| [statusText](#statustext) | string | HTTP ステータステキスト |
| [headers](#headers) | object | レスポンスヘッダー |
| [text()](#text) | string | テキストとしてのレスポンスボディ |
| [json()](#json) | any | JSON ボディをパースした値 |
| [binary()](#binary) | [Buffer](https://nodejs.org/docs/latest-v17.x/api/buffer.html) | バイナリとしてのレスポンスボディ |

## status

レスポンスの HTTP ステータスコード。

### 例

```javascript
const resp = HTTP.get('https://jsonplaceholder.typicode.com/posts/1')
console.log(resp.status) // 200
```

### 戻り値

number - サーバーからの HTTP ステータスコード

## statusText

レスポンスの HTTP ステータステキスト。

### 例

```javascript
const resp = HTTP.get('https://jsonplaceholder.typicode.com/posts/1')
console.log(resp.statusText) // OK
```

### 戻り値

string - サーバーからの HTTP ステータステキスト

## headers

レスポンスヘッダー。

### 例

```javascript
let resp = HTTP.get('https://jsonplaceholder.typicode.com/posts/1')
console.log(resp.headers) // {"content-length":"44","content-type":"text/html; charset=utf-8"}
```

### 戻り値

object - サーバーからのレスポンスヘッダー

## text()

テキストとしてのレスポンスボディ。

### 例

```javascript
let resp = HTTP.get('https://jsonplaceholder.typicode.com/posts/1')
console.log(resp.text()) // this is an example.
```

### 戻り値

string - テキストとして受信したレスポンスボディ

## json()

JSON としてパースしたレスポンスボディ。

### 例

```javascript
let resp = HTTP.get('https://jsonplaceholder.typicode.com/posts/1')
console.log(resp.json()) // {msg:"this is an example."}
```

### 戻り値

Object、Array、string、number、boolean、または null - テキストとして受信したレスポンスボディを `JSON.parse()` した結果

## binary()

バイナリとしてのレスポンスボディ。

### 例

```javascript
let resp = HTTP.get('https://jsonplaceholder.typicode.com/posts/1')
console.log(resp.binary().toString('base64'))
```

### 戻り値

[Buffer](https://nodejs.org/docs/latest-v17.x/api/buffer.html) - Buffer として受信したレスポンスボディ
