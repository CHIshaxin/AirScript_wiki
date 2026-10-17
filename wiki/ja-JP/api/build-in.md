---
title: "概要"
source: https://airsheet.wps.cn/docs/ja-JP/api/build-in.html
updated: 1789700916000
hash: BQuNLPJ3
---

# 概要

組み込み関数は、文字列のエンコード／デコード、暗号化、パラメーターの取得など、開発者がよく行う処理を支援します。

## Crypto

データの暗号化とハッシュ化を行います。

### 例

```javascript
console.log(Crypto.createHash('md5').update('foo').digest('hex')) // acbd18db4cc2f85cedef654fccc4a4d8
console.log(
  Crypto.createHmac('sha256', 'a secret')
    .update('some data to hash')
    .digest('hex')
) // 7fd04df92f636fd450bc841c9418e5825c17f33ad9c87c518115a45971f7f77e
```

### メソッド

| メソッド | 戻り値の型 | 説明 |
| --- | --- | --- |
| createHash(algorithm) | [hash](#hash) | ハッシュ インスタンスを作成します。アルゴリズム: `"md5"`, `"sha1"`, `"sha"`, `"sha256"`, `"sha512"`。 |
| createHmac(algorithm, key) | [hmac](#hmac) | HMAC インスタンスを作成します。サポートされるアルゴリズムは上記と同じです。 |

## hash

[Crypto](#crypto) が生成するハッシュ オブジェクトです。

### メソッド

| メソッド | 戻り値の型 | 説明 |
| --- | --- | --- |
| update(data[, inputEncoding]) | [hash](#hash) | `data` でハッシュを更新します。`data` が文字列で `encoding` が省略されている場合は UTF-8 を使用します。`data` が Buffer の場合、`inputEncoding` は無視されます。繰り返し呼び出せます。 |
| digest([encoding]) | string \| Buffer | ダイジェストを返します。`encoding` を指定すると文字列、指定しないと Buffer を返します。 |

## hmac

[Crypto](#crypto) が生成する HMAC オブジェクトです。

### メソッド

| メソッド | 戻り値の型 | 説明 |
| --- | --- | --- |
| update(data[, inputEncoding]) | [hmac](#hmac) | `data` で HMAC を更新します。`hash.update` と同じルールです。 |
| digest([encoding]) | string \| Buffer | HMAC のダイジェストを返します。 |

## Buffer

`Buffer` インスタンスを作成します。

### 例

```javascript
const buf = Buffer.from([0x62, 0x75, 0x66, 0x66, 0x65, 0x72])
console.log(buf.toString()) // buffer
```

### メソッド

| メソッド | 戻り値の型 | 説明 |
| --- | --- | --- |
| from(array) | Buffer | バイト値 0～255 の配列から新しい Buffer を作成します。 |
| from(string[, encoding]) | Buffer | 文字列から Buffer を作成します。 |
| from(arrayBuffer[, byteOffset[, length]]) | Buffer | `ArrayBuffer` の一部から Buffer を作成します。 |

## Time

スリープを含む時間関連のヘルパーです。

### 例

```javascript
Time.sleep(1000) // 1 秒スリープ
```

### メソッド

| メソッド | 戻り値の型 | 説明 |
| --- | --- | --- |
| sleep(millisecond) | undefined | 指定したミリ秒だけスリープします。 |

## Arguments

設定済みのパラメーター値を読み取ります。

### 例

```javascript
Arguments.get('foo.bar', 'defaults')
// カスタム パラメーターが { foo: { bar: 'value' } } のときは "value" を返します。
// キーがないときは第 2 引数の "defaults" を返します。
```

### メソッド

| メソッド | 戻り値の型 | 説明 |
| --- | --- | --- |
| get(string[, defaults]) | any | ドット区切りのパスでネストしたカスタム パラメーターを読み取ります（例: `a.b` は `obj.a.b`）。キーがないときは `defaults` が指定されていればそれを、なければ `undefined` を返します。 |
