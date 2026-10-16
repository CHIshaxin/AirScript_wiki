---
title: "Network API"
source: https://airsheet.wps.cn/docs/en-US/api/advanced/HTTP.html
updated: 1789700916000
hash: CukGvQzP
---

# Network API

AirScript exposes a global HTTP object. Developers call its methods to request external services; on success the server's response is returned synchronously.

This API behaves much like `fetch()` in the browser, so front-end developers can pick it up quickly.

Tip

Before sending requests with the HTTP object, add the `Network API` service from the script editor toolbar **Services** menu.

### Quick start

```javascript
// Sends a network request
const resp = HTTP.fetch('https://jsonplaceholder.typicode.com/posts/1', {
  timeout: 2000
})
const data = resp.json()
console.log(data.note, data.content)
```

### Methods

| Method | Return type | Summary |
|---|---|---|
| [fetch(url[, options])](#fetch) | [Response](#response) | Sends a custom HTTP request |
| [get(url[, options])](#get) | [Response](#response) | Sends a GET request |
| [delete(url[, options])](#delete) | [Response](#response) | Sends a DELETE request |
| [post(url,body[, options])](#post) | [Response](#response) | Sends a POST request |
| [put(url,body[, options])](#put) | [Response](#response) | Sends a PUT request |

## fetch(url[, options])

Sends an HTTP request; you can set headers and body.

### Example

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

### Parameters

| Name | Type | Default | Required | Description |
|---|---|---|---|---|
| url | string |  | true | URL to request; only hostnames without port numbers are allowed |
| options | [RequestOption](#requestoption) | undefined | false | Optional request settings object, as described below. |

### RequestOption

| Name | Type | Default | Required | Description |
|---|---|---|---|---|
| [method](#method) | string | GET | false | HTTP method, e.g. `GET`, `POST`, `PUT`, `DELETE`, etc. |
| [timeout](#timeout) | number | 10000 | false | Timeout in milliseconds (ms), range 0~60000; values outside the range fall back to 10 seconds. |
| [headers](#headers) | object | undefined | false | Request headers, e.g. `cookie`. |
| [body](#body) | string | undefined | false | Request body. |

### Return value

[Response](#response) - The server's response

## get(url[, options])

Sends a GET request.

### Example

```javascript
const resp = HTTP.get('https://reqres.in/api/users/2')
console.log(resp.json())
```

### Parameters

| Name | Type | Default | Required | Description |
|---|---|---|---|---|
| url | string |  | true | URL to request; only hostnames without port numbers are allowed |
| options | [MethodRequestOption](#methodrequestoption) | undefined | false | Optional settings object for this request, as described below. |

### MethodRequestOption

| Name | Type | Default | Required | Description |
|---|---|---|---|---|
| [timeout](#timeout) | number | 10000 | false | Timeout in milliseconds (ms), range 0~60000; values outside the range fall back to 10 seconds. |
| [headers](#headers) | object | undefined | false | Request headers, e.g. `cookie`. |

### Return value

[Response](#response) - The server's response

## delete(url[, options])

Sends a DELETE request.

### Example

```javascript
const resp = HTTP.delete('https://reqres.in/api/users/2')
console.log(resp.status)
```

### Parameters

| Name | Type | Default | Required | Description |
|---|---|---|---|---|
| url | string |  | true | URL to request; only hostnames without port numbers are allowed |
| options | [MethodRequestOption](#methodrequestoption) | undefined | false | Optional settings object for this request, as described below. |

### Return value

[Response](#response) - The server's response

## post(url,body[, options])

Sends a POST request.

### Example

```javascript
// Sends form
const formResp = HTTP.post(
  'https://www.example.com',
  { foo: 'bar' },
  { headers: { 'content-type': 'multipart/form-data' } }
)

// Sends JSON
const resp = HTTP.post('https://reqres.in/api/users', {
  name: 'morpheus',
  job: 'leader'
})

console.log(resp.json())
```

### Parameters

| Name | Type | Default | Required | Description |
|---|---|---|---|---|
| url | string |  | true | URL to request; only hostnames without port numbers are allowed |
| body | string\| object |  | true | Request body |
| options | [MethodRequestOption](#methodrequestoption) | undefined | false | Optional settings object for this request, as described below. |

### Return value

[Response](#response) - The server's response

## put(url,body[, options])

Sends a PUT request.

### Example

```javascript
const resp = HTTP.put('https://reqres.in/api/users/200', {
  name: 'wps',
  job: 'developer'
})
console.log(resp.json())
```

### Parameters

| Name | Type | Default | Required | Description |
|---|---|---|---|---|
| url | string |  | true | URL to request; only hostnames without port numbers are allowed |
| body | string\| object |  | true | Request body |
| options | [MethodRequestOption](#methodrequestoption) | undefined | false | Optional settings object for this request, as described below. |

### Return value

[Response](#response) - The server's response

## Response

Return value after an HTTP request. The response is stream-like; only the first call to `text()`, `json()`, or `binary()` returns data.

### Example

```javascript
let resp = HTTP.get('https://jsonplaceholder.typicode.com/posts/1')
console.log(resp.status) // 200
console.log(resp.statusText) // OK
console.log(resp.text()) // `{foo:"bar"}`
console.log(resp.json()) // {foo:"bar"}
console.log(resp.status) // [...]
```

### Methods

| Method | Return type | Summary |
|---|---|---|
| [status](#status) | number | HTTP status code |
| [statusText](#statustext) | string | HTTP status text |
| [headers](#status) | object | Response headers |
| [text()](#text) | string | Response body as text |
| [json()](#json) | any | Parses JSON body into a value |
| [binary()](#binary) | [Buffer](https://nodejs.org/docs/latest-v17.x/api/buffer.html) | Response body as binary |

## status

HTTP status code of the response.

### Example

```javascript
const resp = HTTP.get('https://jsonplaceholder.typicode.com/posts/1')
console.log(resp.status) // 200
```

### Return value

number - HTTP status code from the server

## statusText

HTTP status text of the response.

### Example

```javascript
const resp = HTTP.get('https://jsonplaceholder.typicode.com/posts/1')
console.log(resp.statusText) // OK
```

### Return value

string - HTTP status text from the server

## headers

Response headers.

### Example

```javascript
let resp = HTTP.get('https://jsonplaceholder.typicode.com/posts/1')
console.log(resp.headers) // {"content-length":"44","content-type":"text/html; charset=utf-8"}
```

### Return value

object - Response headers from the server

## text()

Response body as text.

### Example

```javascript
let resp = HTTP.get('https://jsonplaceholder.typicode.com/posts/1')
console.log(resp.text()) // this is an example.
```

### Return value

string - Response body received as text

## json()

Response body parsed as JSON.

### Example

```javascript
let resp = HTTP.get('https://jsonplaceholder.typicode.com/posts/1')
console.log(resp.json()) // {msg:"this is an example."}
```

### Return value

Object, Array, string, number, boolean, or null - Response body received as text and returned after `JSON.parse()`

## binary()

Response body as binary.

### Example

```javascript
let resp = HTTP.get('https://jsonplaceholder.typicode.com/posts/1')
console.log(resp.binary().toString('base64'))
```

### Return value

[Buffer](https://nodejs.org/docs/latest-v17.x/api/buffer.html) - Response body received as a Buffer
