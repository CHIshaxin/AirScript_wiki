---
title: "概述"
source: https://airsheet.wps.cn/docs/api/build-in.html
updated: 1789700916000
hash: XqBgjGeS
---

# 概述

内置函数是用来帮助开发者处理字符串编码/解码、信息处理、参数获取和其他杂项任务的实用函数

## Crypto

对信息进行加密，摘要处理

### 示例

```javascript
// 摘要foo这个字符串信息
console.log(Crypto.createHash("md5").update("foo").digest("hex")) // acbd18db4cc2f85cedef654fccc4a4d8
console.log(Crypto.createHmac("sha256", "a secret").update('some data to hash').digest('hex')) //7fd04df92f636fd450bc841c9418e5825c17f33ad9c87c518115a45971f7f77e
```

### 方法列表

| 方法 | 返回类型 | 简介 |
| --- | --- | --- |
| createHash(algorithm) | [hash](#hash) | 创建摘要算法实例，允许"md5", "sha1", "sha", "sha256", "sha512" |
| createHmac(algorithm, key) | [hmac](#hmac) | 创建HMAC算法实例，允许"md5", "sha1", "sha", "sha256", "sha512" |

## hash

摘要对象，由[Crypto](#crypto)产生

### 方法列表

| 方法 | 返回类型 | 简介 |
| --- | --- | --- |
| update(data[ ,inputEncoding]) | [hash](#hash) | 使用给定的 data 更新哈希内容，如果未提供 encoding，且 data 是字符串，则强制为 'utf8' 编码，如果 data 是 Buffer,则忽略 inputEncoding,可重复调用添加数据 |
| digest([encoding] | string\| Buffer | 计算传给被哈希的所有数据的摘要，如果提供了 encoding，则将返回字符串；否则返回 Buffer。 |

## hmac

hmac对象，由[Crypto](#crypto)产生

### 方法列表

| 方法 | 返回类型 | 简介 |
| --- | --- | --- |
| update(data[ ,inputEncoding]) | [hash](#hash) | 使用给定的 data 更新hmac内容，如果未提供 encoding，且 data 是字符串，则强制为 'utf8' 编码，如果 data 是 Buffer,则忽略 inputEncoding,可重复调用添加数据 |
| digest([encoding] | string\| Buffer | 计算传给被hmac的所有数据的摘要，如果提供了 encoding，则将返回字符串；否则返回 Buffer。 |

## Buffer

产生一个 Buffer 实例

### 示例

```javascript
// 创建包含字符串 'buffer' 的 UTF-8 字节的新缓冲区。
const buf = Buffer.from([0x62, 0x75, 0x66, 0x66, 0x65, 0x72]);
console.log(buf.toString()) // buffer
```

### 方法列表

| 方法 | 返回类型 | 简介 |
| --- | --- | --- |
| from(array) | Buffer | 使用 0 – 255 范围内的字节 array 分配新的 Buffer。 |
| from(string[, encoding]) | Buffer | 从字符串转化为Buffer |
| from(arrayBuffer[, byteOffset[, length]]) | Buffer | 截断arrayBuffer的部分字节，生成新的Buffer |

## Time

时间函数，提供如休眠的方法

### 示例

```javascript
Time.sleep(1000) // 休眠一秒
```

### 方法列表

| 方法 | 返回类型 | 简介 |
| --- | --- | --- |
| sleep(millisecond) | undefined | 休眠指定毫秒数 |

## Arguments

方便获取配置的参数数据

### 示例

```javascript
Arguments.get("foo.bar", "defaults") // 如果自定义参数是{foo : {bar : "value"}}，则返回"value"，如果不存在，则返回第二个参数"defaults"
```

### 方法列表

| 方法 | 返回类型 | 简介 |
| --- | --- | --- |
| get(string[, defaults]) | any | 通过获取自定义参数的值，key支持使用.进行多次查找，如a.b会寻找{a : {b : ""}}这个结构体的b值。可指定默认值，如果找不到key对应的自定义参数，就返回默认值，没有指定默认值也找不到key返回undefined |
