---
title: "Overview"
source: https://airsheet.wps.cn/docs/en-US/api/build-in.html
updated: 1789700916000
hash: qlIgmP6g
---

# Overview

Built-in functions help developers with string encoding and decoding, cryptography, parameter access, and other common tasks.

## Crypto

Encrypt and hash data.

### Example

```javascript
console.log(Crypto.createHash('md5').update('foo').digest('hex')) // acbd18db4cc2f85cedef654fccc4a4d8
console.log(
  Crypto.createHmac('sha256', 'a secret')
    .update('some data to hash')
    .digest('hex')
) // 7fd04df92f636fd450bc841c9418e5825c17f33ad9c87c518115a45971f7f77e
```

### Methods

| Method | Return type | Description |
| --- | --- | --- |
| createHash(algorithm) | [hash](#hash) | Create a hash instance. Algorithms: `"md5"`, `"sha1"`, `"sha"`, `"sha256"`, `"sha512"`. |
| createHmac(algorithm, key) | [hmac](#hmac) | Create an HMAC instance. Same supported algorithms. |

## hash

Hash object produced by [Crypto](#crypto).

### Methods

| Method | Return type | Description |
| --- | --- | --- |
| update(data[, inputEncoding]) | [hash](#hash) | Update the hash with `data`. If `data` is a string and `encoding` is omitted, UTF-8 is used. If `data` is a Buffer, `inputEncoding` is ignored. Can be called repeatedly. |
| digest([encoding]) | string \| Buffer | Returns the digest. If `encoding` is set, returns a string; otherwise a Buffer. |

## hmac

HMAC object produced by [Crypto](#crypto).

### Methods

| Method | Return type | Description |
| --- | --- | --- |
| update(data[, inputEncoding]) | [hmac](#hmac) | Update the HMAC with `data`. Same rules as `hash.update`. |
| digest([encoding]) | string \| Buffer | Returns the HMAC digest. |

## Buffer

Creates a `Buffer` instance.

### Example

```javascript
const buf = Buffer.from([0x62, 0x75, 0x66, 0x66, 0x65, 0x72])
console.log(buf.toString()) // buffer
```

### Methods

| Method | Return type | Description |
| --- | --- | --- |
| from(array) | Buffer | New Buffer from byte values 0–255. |
| from(string[, encoding]) | Buffer | Buffer from a string. |
| from(arrayBuffer[, byteOffset[, length]]) | Buffer | Buffer from part of an `ArrayBuffer`. |

## Time

Time helpers, including sleep.

### Example

```javascript
Time.sleep(1000) // sleep one second
```

### Methods

| Method | Return type | Description |
| --- | --- | --- |
| sleep(millisecond) | undefined | Sleep for the given number of milliseconds. |

## Arguments

Read configured parameter values.

### Example

```javascript
Arguments.get('foo.bar', 'defaults')
// If custom params are { foo: { bar: 'value' } }, returns "value";
// if missing, returns the second argument "defaults".
```

### Methods

| Method | Return type | Description |
| --- | --- | --- |
| get(string[, defaults]) | any | Read a nested custom parameter using dot paths (e.g. `a.b` reads `obj.a.b`). If the key is missing, returns `defaults` when provided, otherwise `undefined`. |
