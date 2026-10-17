---
title: "网络 API"
source: https://airsheet.wps.cn/docs/apiV2/advanced/HTTP.html
updated: 1789700916000
hash: O-Vl9cpL
---

# 网络 API

AirScript 提供一个全局的 HTTP 对象，开发者可通过此对象提供的方法请求外部服务，请求成功后会同步返回服务器的响应。

该 API 的使用方式与浏览器内的 fetch()函数基本一致，对于前端开发者来说应该可以很快上手。

提示

在使用 HTTP 对象提供的方法发送请求之前，确保您已添加`网络API`服务，在脚本编辑器的【工具栏】-【服务】菜单内添加即可。

### 快速使用

```javascript
// 发起网络请求
const resp = HTTP.fetch('https://open.iciba.com/dsapi/', {
  timeout: 2000
})
const data = resp.json()
console.log(data.note, data.content)
```

### 方法列表

| 方法 | 返回类型 | 简介 |
| --- | --- | --- |
| [fetch(url[, options])](#fetch) | [Response](#response) | 发起自定义类型的网络请求 |
| [get(url[, options])](#get) | [Response](#response) | 发起 GET 类型的网络请求 |
| [delete(url[, options])](#delete) | [Response](#response) | 发起 DELETE 类型的网络请求 |
| [post(url,body[, options])](#post) | [Response](#response) | 发起 POST 类型的网络请求 |
| [put(url,body[, options])](#put) | [Response](#response) | 发起 PUT 类型的网络请求 |

## fetch(url[, options])

发起一个网络请求，可以自定义设置 headers 和 body。

### 示例

```javascript
const resp = HTTP.fetch('https://www.kdocs.cn', {
  method: 'GET',
  timeout: 2000,
  headers: {
    'User-Agent':
      'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/109.0.0.0 Safari/537.36'
  }
})
console.log(resp.text())
```

### 参数

| 名称 | 类型 | 默认值 | 必填项 | 说明 |
| --- | --- | --- | --- | --- |
| url | string |  | true | 需要访问的网络地址，只允许访问不带端口号的域名 |
| options | [RequestOption](#requestoption) | undefined | false | 一个 JavaScript 对象，可指定发起请求的可选参数，如下所示。 |

### RequestOption

| 名称 | 类型 | 默认值 | 必填项 | 说明 |
| --- | --- | --- | --- | --- |
| [method](#method) | string | GET | false | 发起网络请求的方法，例如`GET`、`POST`、`PUT`、`DELETE`等 |
| [timeout](#timeout) | number | 10000 | false | 发起网络请求的超时时间，单位毫秒(ms)，数据范围为 0~60000，超出范围的数据将被设为默认值 10 秒。 |
| [headers](#headers) | object | undefined | false | 发起网络请求的头部。例如`cookie`等 |
| [body](#body) | string | undefined | false | 发起网络请求的主体内容。 |

### 返回值

[Response](#response) - 服务器返回的响应

## get(url[, options])

发起 GET 类型的网络请求。

### 示例

```javascript
const resp = HTTP.get('https://reqres.in/api/users/2')
console.log(resp.json())
```

### 参数

| 名称 | 类型 | 默认值 | 必填项 | 说明 |
| --- | --- | --- | --- | --- |
| url | string |  | true | 需要访问的网络地址，只允许访问不带端口号的域名 |
| options | [MethodRequestOption](#methodrequestoption) | undefined | false | 一个 JavaScript 对象，可指定特定请求的可选参数，如下所示。 |

### MethodRequestOption

| 名称 | 类型 | 默认值 | 必填项 | 说明 |
| --- | --- | --- | --- | --- |
| [timeout](#timeout) | number | 10000 | false | 发起网络请求的超时时间，单位毫秒(ms)，数据范围为 0~60000，超出范围的数据将被设为默认值 10 秒。 |
| [headers](#headers) | object | undefined | false | 发起网络请求的头部。例如`cookie`等 |

### 返回值

[Response](#response) - 服务器返回的响应

## delete(url[, options])

发起 DELETE 类型的网络请求。

### 示例

```javascript
const resp = HTTP.delete('https://reqres.in/api/users/2')
console.log(resp.status)
```

### 参数

| 名称 | 类型 | 默认值 | 必填项 | 说明 |
| --- | --- | --- | --- | --- |
| url | string |  | true | 需要访问的网络地址，只允许访问不带端口号的域名 |
| options | [MethodRequestOption](#methodrequestoption) | undefined | false | 一个 JavaScript 对象，可指定特定请求的可选参数，如下所示。 |

### 返回值

[Response](#response) - 服务器返回的响应

## post(url,body[, options])

发起 POST 类型的网络请求。

### 示例

```javascript
// 发送form
const formResp = HTTP.post(
  'https://www.example.cn',
  { foo: 'bar' },
  { headers: { 'content-type': 'multipart/form-data' } }
)

//发送json
const resp = HTTP.post('https://reqres.in/api/users', {
  name: 'morpheus',
  job: 'leader'
})

console.log(resp.json())
```

### 参数

| 名称 | 类型 | 默认值 | 必填项 | 说明 |
| --- | --- | --- | --- | --- |
| url | string |  | true | 需要访问的网络地址，只允许访问不带端口号的域名 |
| body | string\| object |  | true | 请求体 |
| options | [MethodRequestOption](#methodrequestoption) | undefined | false | 一个 JavaScript 对象，可指定特定请求的可选参数，如下所示。 |

### 返回值

[Response](#response) - 服务器返回的响应

## put(url,body[, options])

发起 PUT 类型的网络请求。

### 示例

```javascript
const resp = HTTP.put('https://reqres.in/api/users/200', {
  name: 'wps',
  job: 'developer'
})
console.log(resp.json())
```

### 参数

| 名称 | 类型 | 默认值 | 必填项 | 说明 |
| --- | --- | --- | --- | --- |
| url | string |  | true | 需要访问的网络地址，只允许访问不带端口号的域名 |
| body | string\| object |  | true | 请求体 |
| options | [MethodRequestOption](#methodrequestoption) | undefined | false | 一个 JavaScript 对象，可指定特定请求的可选参数，如下所示。 |

### 返回值

[Response](#response) - 服务器返回的响应

## Response

HTTP 发起网络请求后返回的响应，response 是流数据，只有首次调用 text()，json()或 binary()能获取到数据

### 示例

```javascript
let resp = HTTP.get('https://open.iciba.com/dsapi/')
console.log(resp.status) // 200
console.log(resp.statusText) // OK
console.log(resp.text()) // `{foo:"bar"}`
console.log(resp.json()) // {foo:"bar"}
console.log(resp.status) // [...]
```

### 方法列表

| 方法 | 返回类型 | 简介 |
| --- | --- | --- |
| [status](#status) | number | 获取响应的 HTTP 状态码 |
| [statusText](#statustext) | string | 获取响应的 HTTP 状态 |
| [headers](#status) | object | 获取响应的 header |
| [text()](#text) | string | 获取服务器返回的文本 Body |
| [json()](#json) | any | 将服务器返回的 json 类型的 Body 转化为结构体 |
| [binary()](#binary) | [Buffer](https://nodejs.org/docs/latest-v17.x/api/buffer.html) | 获取服务器返回的二进制结构的 Body |

## status

获取响应的 HTTP 状态码

### 示例

```javascript
const resp = HTTP.get('https://open.iciba.com/dsapi/')
console.log(resp.status) // 200
```

### 返回值

number - 服务器返回响应的 HTTP 状态码

## statusText

获取响应的 HTTP 状态

### 示例

```javascript
const resp = HTTP.get('https://open.iciba.com/dsapi/')
console.log(resp.statusText) // OK
```

### 返回值

string - 服务器返回响应的 HTTP 状态

## headers

获取响应的 header

### 示例

```javascript
let resp = HTTP.get('https://open.iciba.com/dsapi/')
console.log(resp.headers) // {"content-length":"44","content-type":"text/html; charset=utf-8"}
```

### 返回值

object - 服务器返回响应的 header

## text()

获取服务器返回的 Body

### 示例

```javascript
let resp = HTTP.get('https://open.iciba.com/dsapi/')
console.log(resp.text()) // this is an example.
```

### 返回值

string - 服务器返回的响应的 Body，以文本接受并返回

## json()

获取服务器返回的 Body

### 示例

```javascript
let resp = HTTP.get('https://open.iciba.com/dsapi/')
console.log(resp.json()) // {msg:"this is an example."}
```

### 返回值

Object, Array, string, number, boolean, or null - 服务器返回的响应的 Body，以文本接受并经过 JSON.parse()后返回

## binary()

获取服务器返回的 Body

### 示例

```javascript
let resp = HTTP.get('https://open.iciba.com/dsapi/')
console.log(resp.binary().toString('base64'))
```

### 返回值

[Buffer](https://nodejs.org/docs/latest-v17.x/api/buffer.html) - 服务器返回的响应的 Body，以 Buffer 接受二进制数据并返回
