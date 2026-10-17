---
title: "ネットワーク API"
source: https://airsheet.wps.cn/docs/ja-JP/apiV2/advanced/HTTP.html
updated: 1789700916000
hash: DKV7V1q_
---

# ネットワーク API

AirScript はグローバルな HTTP オブジェクトを提供します。開発者はこのオブジェクトのメソッドを使って外部サービスへリクエストでき、成功時にはサーバーのレスポンスが同期的に返ります。

この API の使い方はブラウザの fetch() 関数とほぼ同じため、フロントエンド開発者はすぐに慣れることができます。

ヒント

HTTP オブジェクトのメソッドでリクエストを送る前に、`ネットワーク API` サービスを追加済みであることを確認してください。スクリプトエディターの［ツールバー］－［サービス］メニューから追加できます。

### クイックスタート

```javascript
// ネットワークリクエストを送信
const resp = HTTP.fetch('https://jsonplaceholder.typicode.com/posts/1', {
  timeout: 2000
})
const data = resp.json()
console.log(data.note, data.content)
```

### メソッド一覧

| メソッド | 戻り値の型 | 説明 |
| --- | --- | --- |
| [fetch(url[, options])](#fetch) | [Response](#response) | カスタムタイプのネットワークリクエストを送信 |
| [get(url[, options])](#get) | [Response](#response) | GET タイプのネットワークリクエストを送信 |
| [delete(url[, options])](#delete) | [Response](#response) | DELETE タイプのネットワークリクエストを送信 |
| [post(url,body[, options])](#post) | [Response](#response) | POST タイプのネットワークリクエストを送信 |
| [put(url,body[, options])](#put) | [Response](#response) | PUT タイプのネットワークリクエストを送信 |

## fetch(url[, options])

カスタムヘッダーとボディ設定でネットワークリクエストを送信します。

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
| url | string |  | はい | アクセスするネットワークアドレス。ポート番号を含まないドメイン名のみ許可 |
| options | [RequestOption](#requestoption) | undefined | いいえ | リクエストの任意パラメーターを指定できる JavaScript オブジェクト。下記参照。 |

### RequestOption

| パラメーター | 型 | 既定値 | 必須 | 説明 |
| --- | --- | --- | --- | --- |
| [method](#method) | string | GET | いいえ | ネットワークリクエストのメソッド。`GET`、`POST`、`PUT`、`DELETE` など。 |
| [timeout](#timeout) | number | 10000 | いいえ | ネットワークリクエストのタイムアウト（ミリ秒）。範囲は 0～60000。範囲外は既定で 10 秒。 |
| [headers](#headers) | object | undefined | いいえ | ネットワークリクエストのヘッダー。`cookie` など。 |
| [body](#body) | string | undefined | いいえ | ネットワークリクエストのボディ。 |

### 戻り値

[Response](#response) — サーバーから返されたレスポンス

## get(url[, options])

GET タイプのネットワークリクエストを送信します。

### 例

```javascript
const resp = HTTP.get('https://reqres.in/api/users/2')
console.log(resp.json())
```

### パラメーター

| パラメーター | 型 | 既定値 | 必須 | 説明 |
| --- | --- | --- | --- | --- |
| url | string |  | はい | アクセスするネットワークアドレス。ポート番号を含まないドメイン名のみ許可 |
| options | [MethodRequestOption](#methodrequestoption) | undefined | いいえ | リクエストの任意パラメーターを指定できる JavaScript オブジェクト。下記参照。 |

### MethodRequestOption

| パラメーター | 型 | 既定値 | 必須 | 説明 |
| --- | --- | --- | --- | --- |
| [timeout](#timeout) | number | 10000 | いいえ | ネットワークリクエストのタイムアウト（ミリ秒）。範囲は 0～60000。範囲外は既定で 10 秒。 |
| [headers](#headers) | object | undefined | いいえ | ネットワークリクエストのヘッダー。`cookie` など。 |

### 戻り値

[Response](#response) — サーバーから返されたレスポンス

## delete(url[, options])

DELETE タイプのネットワークリクエストを送信します。

### 例

```javascript
const resp = HTTP.delete('https://reqres.in/api/users/2')
console.log(resp.status)
```

### パラメーター

| パラメーター | 型 | 既定値 | 必須 | 説明 |
| --- | --- | --- | --- | --- |
| url | string |  | はい | アクセスするネットワークアドレス。ポート番号を含まないドメイン名のみ許可 |
| options | [MethodRequestOption](#methodrequestoption) | undefined | いいえ | リクエストの任意パラメーターを指定できる JavaScript オブジェクト。下記参照。 |

### 戻り値

[Response](#response) — サーバーから返されたレスポンス

## post(url,body[, options])

POST タイプのネットワークリクエストを送信します。

### 例

```javascript
// フォームを送信
const formResp = HTTP.post(
  'https://www.example.com',
  { foo: 'bar' },
  { headers: { 'content-type': 'multipart/form-data' } }
)

// JSON を送信
const resp = HTTP.post('https://reqres.in/api/users', {
  name: 'morpheus',
  job: 'leader'
})

console.log(resp.json())
```

### パラメーター

| パラメーター | 型 | 既定値 | 必須 | 説明 |
| --- | --- | --- | --- | --- |
| url | string |  | はい | アクセスするネットワークアドレス。ポート番号を含まないドメイン名のみ許可 |
| body | string\| object |  | はい | リクエストボディ |
| options | [MethodRequestOption](#methodrequestoption) | undefined | いいえ | リクエストの任意パラメーターを指定できる JavaScript オブジェクト。下記参照。 |

### 戻り値

[Response](#response) — サーバーから返されたレスポンス

## put(url,body[, options])

PUT タイプのネットワークリクエストを送信します。

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
| url | string |  | はい | アクセスするネットワークアドレス。ポート番号を含まないドメイン名のみ許可 |
| body | string\| object |  | はい | リクエストボディ |
| options | [MethodRequestOption](#methodrequestoption) | undefined | いいえ | リクエストの任意パラメーターを指定できる JavaScript オブジェクト。下記参照。 |

### 戻り値

[Response](#response) — サーバーから返されたレスポンス

## Response

HTTP がネットワークリクエストを送信した後に返るレスポンス。Response はストリームデータであり、text()、json()、binary() のいずれかを初めて呼び出したときにのみデータを取得できます。

### 例

```javascript
let resp = HTTP.get('https://jsonplaceholder.typicode.com/posts/1')
console.log(resp.status) // 200
console.log(resp.statusText) // OK
console.log(resp.text()) // `{foo:"bar"}`
console.log(resp.json()) // {foo:"bar"}
console.log(resp.status) // [...]
```

### メソッド一覧

| メソッド | 戻り値の型 | 説明 |
| --- | --- | --- |
| [status](#status) | number | レスポンスの HTTP ステータスコードを取得 |
| [statusText](#statustext) | string | レスポンスの HTTP ステータスを取得 |
| [headers](#status) | object | レスポンスヘッダーを取得 |
| [text()](#text) | string | サーバーが返したテキスト Body を取得 |
| [json()](#json) | any | サーバーが返した JSON 型の Body をオブジェクトに変換 |
| [binary()](#binary) | [Buffer](https://nodejs.org/docs/latest-v17.x/api/buffer.html) | サーバーが返したバイナリ Body を取得 |

## status

レスポンスの HTTP ステータスコードを取得します。

### 例

```javascript
const resp = HTTP.get('https://jsonplaceholder.typicode.com/posts/1')
console.log(resp.status) // 200
```

### 戻り値

number — サーバーレスポンスの HTTP ステータスコード

## statusText

レスポンスの HTTP ステータスを取得します。

### 例

```javascript
const resp = HTTP.get('https://jsonplaceholder.typicode.com/posts/1')
console.log(resp.statusText) // OK
```

### 戻り値

string — サーバーレスポンスの HTTP ステータス

## headers

レスポンスヘッダーを取得します。

### 例

```javascript
let resp = HTTP.get('https://jsonplaceholder.typicode.com/posts/1')
console.log(resp.headers) // {"content-length":"44","content-type":"text/html; charset=utf-8"}
```

### 戻り値

object — サーバーレスポンスのヘッダー

## text()

サーバーが返した Body を取得します。

### 例

```javascript
let resp = HTTP.get('https://jsonplaceholder.typicode.com/posts/1')
console.log(resp.text()) // this is an example.
```

### 戻り値

string — サーバーレスポンスの Body。テキストとして受信して返す

## json()

サーバーが返した Body を取得します。

### 例

```javascript
let resp = HTTP.get('https://jsonplaceholder.typicode.com/posts/1')
console.log(resp.json()) // {msg:"this is an example."}
```

### 戻り値

Object、Array、string、number、boolean、または null — サーバーレスポンスの Body。テキストとして受信し、JSON.parse() 後に返す

## binary()

サーバーが返した Body を取得します。

### 例

```javascript
let resp = HTTP.get('https://jsonplaceholder.typicode.com/posts/1')
console.log(resp.binary().toString('base64'))
```

### 戻り値

[Buffer](https://nodejs.org/docs/latest-v17.x/api/buffer.html) — サーバーレスポンスの Body。バイナリデータとして Buffer で受信して返す
