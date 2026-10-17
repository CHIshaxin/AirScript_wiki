---
title: "Network API"
source: https://airsheet.wps.cn/docs/en-US/apiV2/advanced/HTTP.html
updated: 1789700916000
hash: D2PZKARN
---

# Network API

AirScript provides a global HTTP object. Developers can use the methods provided by this object to request external services, and the server's response will be returned synchronously after a successful request.

The usage of this API is basically the same as the fetch() function in browsers, so frontend developers should be able to get started quickly.

Tip

Before using the methods provided by the HTTP object to send requests, make sure you have added the `Network API` service. You can add it in the [Toolbar] - [Services] menu of the script editor.

### Quick Start

```javascript
// Send network request
const resp = HTTP.fetch('https://jsonplaceholder.typicode.com/posts/1', {
  timeout: 2000
})
const data = resp.json()
console.log(data.note, data.content)
```

### Method List

| Method | Return Type | Description |
| --- | --- | --- |
| [fetch(url[, options])](#fetch) | [Response](#response) | Send a custom type network request |
| [get(url[, options])](#get) | [Response](#response) | Send a GET type network request |
| [delete(url[, options])](#delete) | [Response](#response) | Send a DELETE type network request |
| [post(url,body[, options])](#post) | [Response](#response) | Send a POST type network request |
| [put(url,body[, options])](#put) | [Response](#response) | Send a PUT type network request |

## fetch(url[, options])

Send a network request with custom headers and body settings.

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
| --- | --- | --- | --- | --- |
| url | string |  | true | Network address to access, only domain names without port numbers are allowed |
| options | [RequestOption](#requestoption) | undefined | false | A JavaScript object that can specify optional parameters for the request, as shown below. |

### RequestOption

| Name | Type | Default | Required | Description |
| --- | --- | --- | --- | --- |
| [method](#method) | string | GET | false | Method for network request, such as `GET`, `POST`, `PUT`, `DELETE`, etc. |
| [timeout](#timeout) | number | 10000 | false | Timeout for network request in milliseconds (ms), range is 0~60000, values outside range default to 10 seconds. |
| [headers](#headers) | object | undefined | false | Headers for network request, such as `cookie`, etc. |
| [body](#body) | string | undefined | false | Body content for network request. |

### Return Value

[Response](#response) - Response returned by the server

## get(url[, options])

Send a GET type network request.

### Example

```javascript
const resp = HTTP.get('https://reqres.in/api/users/2')
console.log(resp.json())
```

### Parameters

| Name | Type | Default | Required | Description |
| --- | --- | --- | --- | --- |
| url | string |  | true | Network address to access, only domain names without port numbers are allowed |
| options | [MethodRequestOption](#methodrequestoption) | undefined | false | A JavaScript object that can specify optional parameters for the request, as shown below. |

### MethodRequestOption

| Name | Type | Default | Required | Description |
| --- | --- | --- | --- | --- |
| [timeout](#timeout) | number | 10000 | false | Timeout for network request in milliseconds (ms), range is 0~60000, values outside range default to 10 seconds. |
| [headers](#headers) | object | undefined | false | Headers for network request, such as `cookie`, etc. |

### Return Value

[Response](#response) - Response returned by the server

## delete(url[, options])

Send a DELETE type network request.

### Example

```javascript
const resp = HTTP.delete('https://reqres.in/api/users/2')
console.log(resp.status)
```

### Parameters

| Name | Type | Default | Required | Description |
| --- | --- | --- | --- | --- |
| url | string |  | true | Network address to access, only domain names without port numbers are allowed |
| options | [MethodRequestOption](#methodrequestoption) | undefined | false | A JavaScript object that can specify optional parameters for the request, as shown below. |

### Return Value

[Response](#response) - Response returned by the server

## post(url,body[, options])

Send a POST type network request.

### Example

```javascript
// Send form
const formResp = HTTP.post(
  'https://www.example.com',
  { foo: 'bar' },
  { headers: { 'content-type': 'multipart/form-data' } }
)

// Send json
const resp = HTTP.post('https://reqres.in/api/users', {
  name: 'morpheus',
  job: 'leader'
})

console.log(resp.json())
```

### Parameters

| Name | Type | Default | Required | Description |
| --- | --- | --- | --- | --- |
| url | string |  | true | Network address to access, only domain names without port numbers are allowed |
| body | string\| object |  | true | Request body |
| options | [MethodRequestOption](#methodrequestoption) | undefined | false | A JavaScript object that can specify optional parameters for the request, as shown below. |

### Return Value

[Response](#response) - Response returned by the server

## put(url,body[, options])

Send a PUT type network request.

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
| --- | --- | --- | --- | --- |
| url | string |  | true | Network address to access, only domain names without port numbers are allowed |
| body | string\| object |  | true | Request body |
| options | [MethodRequestOption](#methodrequestoption) | undefined | false | A JavaScript object that can specify optional parameters for the request, as shown below. |

### Return Value

[Response](#response) - Response returned by the server

## Response

Response returned after HTTP sends a network request. Response is stream data, and data can only be obtained on the first call to text(), json(), or binary().

### Example

```javascript
let resp = HTTP.get('https://jsonplaceholder.typicode.com/posts/1')
console.log(resp.status) // 200
console.log(resp.statusText) // OK
console.log(resp.text()) // `{foo:"bar"}`
console.log(resp.json()) // {foo:"bar"}
console.log(resp.status) // [...]
```

### Method List

| Method | Return Type | Description |
| --- | --- | --- |
| [status](#status) | number | Get response HTTP status code |
| [statusText](#statustext) | string | Get response HTTP status |
| [headers](#status) | object | Get response header |
| [text()](#text) | string | Get text Body returned by server |
| [json()](#json) | any | Convert json type Body returned by server to object |
| [binary()](#binary) | [Buffer](https://nodejs.org/docs/latest-v17.x/api/buffer.html) | Get binary Body returned by server |

## status

Get response HTTP status code

### Example

```javascript
const resp = HTTP.get('https://jsonplaceholder.typicode.com/posts/1')
console.log(resp.status) // 200
```

### Return Value

number - HTTP status code of server response

## statusText

Get response HTTP status

### Example

```javascript
const resp = HTTP.get('https://jsonplaceholder.typicode.com/posts/1')
console.log(resp.statusText) // OK
```

### Return Value

string - HTTP status of server response

## headers

Get response header

### Example

```javascript
let resp = HTTP.get('https://jsonplaceholder.typicode.com/posts/1')
console.log(resp.headers) // {"content-length":"44","content-type":"text/html; charset=utf-8"}
```

### Return Value

object - Header of server response

## text()

Get Body returned by server

### Example

```javascript
let resp = HTTP.get('https://jsonplaceholder.typicode.com/posts/1')
console.log(resp.text()) // this is an example.
```

### Return Value

string - Body of server response, received and returned as text

## json()

Get Body returned by server

### Example

```javascript
let resp = HTTP.get('https://jsonplaceholder.typicode.com/posts/1')
console.log(resp.json()) // {msg:"this is an example."}
```

### Return Value

Object, Array, string, number, boolean, or null - Body of server response, received as text and returned after JSON.parse()

## binary()

Get Body returned by server

### Example

```javascript
let resp = HTTP.get('https://jsonplaceholder.typicode.com/posts/1')
console.log(resp.binary().toString('base64'))
```

### Return Value

[Buffer](https://nodejs.org/docs/latest-v17.x/api/buffer.html) - Body of server response, received as Buffer for binary data and returned
