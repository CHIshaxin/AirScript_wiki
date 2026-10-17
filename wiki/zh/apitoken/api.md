---
title: "接口说明"
source: https://airsheet.wps.cn/docs/apitoken/api.html
updated: 1789700916000
hash: B6NYn_4E
---

# 接口说明

成功生成脚本令牌后，就可以通过 HTTP 接口执行脚本了，我们提供了**同步执行**和**异步执行**两种脚本执行接口供开发者使用。

相较而言，前者使用更简单，接口调用后会直接返回执行结果，适用于执行耗时一般的场景；而后者则略微复杂一点，接口调用后不会返回最终的执行的结果，但会立即返回一个`task_id`，您需要根据此`task_id`轮询脚本执行的日志，而无需同步等待结果阻塞业务流程，该接口适用于执行耗时比较大的场景。

无论使用您使用哪个接口，都必须先获取到文件 ID 和脚本 ID，请先进入脚本编辑器，在侧边栏列表的更多菜单里复制 webhook 链接即可。

## 同步执行脚本

`POST /api/v3/ide/file/:file_id/script/:script_id/sync_task`

### Header 参数

| 参数 | 必须 | 类型 | 说明 |
| --- | --- | --- | --- |
| Content-Type | 是 | string | `application/json` |
| AirScript-Token | 是 | string | 传入您通过 AirScript 编辑器生成的**脚本令牌**（APIToken） |

### path 参数

| 参数 | 必须 | 类型 | 说明 |
| --- | --- | --- | --- |
| script_id | 是 | string | 脚本的 ID |
| file_id | 是 | string | 运行脚本的文件 ID |

### body 参数

| 参数 | 必须 | 类型 | 说明 |
| --- | --- | --- | --- |
| Context | 是 | Object | 运行时的上下文参数 |
| Context.argv | 否 | Object | 传入的上下文参数对象，比如传入`{name: 'xiaomeng', age: 18}`，在 AS 代码中可通过`Context.argv.name`获取到传入的值 |
| Context.sheet_name | 否 | string | et,ksheet 运行时所在表名 |
| Context.range | 否 | string | et,ksheet 运行时所在区域，例如`$B$156` |
| Context.link_from | 否 | string | et,ksheet 点击超链接所在单元格 |
| Context.db_active_view | 否 | string | db 运行时所在 view 名 |
| Context.db_selection | 否 | string | db 运行时所在选区 |

### 返回参数

| 参数 | 必须 | 类型 | 说明 |
| --- | --- | --- | --- |
| data | 是 | Object | 任务执行数据对象 |
| data.result | 是 | string | 任务执行返回的数据 |
| data.logs | 是 | Array | 任务执行日志 |
| data.logs[i].filename | 是 | string | 执行文件的名称 |
| data.logs[i].timestamp | 是 | string | 执行时间 |
| data.logs[i].unix_time | 是 | number | 执行 unix 时间戳 |
| data.logs[i].level | 是 | string | 日志级别 |
| data.logs[i].args | 是 | string[] | 日志打印参数 |
| status | 是 | string | 任务是否执行完毕 |
| error | 是 | string | 任务执行错误信息 |
| error_details | 否 | object | 错误信息详情对象 |
| error_details.name | 否 | string | 错误信息名称 |
| error_details.msg | 否 | string | 错误信息 |
| error_details.stack | 否 | string[] | 错误信息栈 |
| error_details.unix_time | 否 | number | 错误信息 unix 时间 |

### 请求示例

Shell

```shell
curl --request POST \
	--url https://www.kdocs.cn/api/v3/ide/file/:file_id/script/:script_id/sync_task \
	--header 'AirScript-Token: xxx' \
	--header 'Content-Type: application/json' \
	--data '{"Context":{"argv":{},"sheet_name":"表名","range":"$B$156"}}'
```

Java

```java
OkHttpClient client = new OkHttpClient();

MediaType mediaType = MediaType.parse("application/json");
RequestBody body = RequestBody.create(mediaType, "{\"Context\":{\"argv\":{},\"sheet_name\":\"表名\",\"range\":\"$B$156\"}}");
Request request = new Request.Builder()
	.url("https://www.kdocs.cn/api/v3/ide/file/:file_id/script/:script_id/sync_task")
	.post(body)
	.addHeader("Content-Type", "application/json")
	.addHeader("AirScript-Token", "xxx")
	.build();

Response response = client.newCall(request).execute();
```

Go

```go
package main

import (
	"fmt"
	"strings"
	"net/http"
	"io/ioutil"
)

func main() {

	url := "https://www.kdocs.cn/api/v3/ide/file/:file_id/script/:script_id/sync_task"

	payload := strings.NewReader("{\"Context\":{\"argv\":{},\"sheet_name\":\"表名\",\"range\":\"$B$156\"}}")

	req, _ := http.NewRequest("POST", url, payload)

	req.Header.Add("Content-Type", "application/json")
	req.Header.Add("AirScript-Token", "xxx")

	res, _ := http.DefaultClient.Do(req)

	defer res.Body.Close()
	body, _ := ioutil.ReadAll(res.Body)

	fmt.Println(res)
	fmt.Println(string(body))

}
```

Python

```python
import http.client

conn = http.client.HTTPSConnection("www.kdocs.cn")

payload = "{\"Context\":{\"argv\":{},\"sheet_name\":\"表名\",\"range\":\"$B$156\"}}"

headers = {
    'Content-Type': "application/json",
    'AirScript-Token': "xxx"
    }

conn.request("POST", "/api/v3/ide/file/:file_id/script/:script_id/sync_task", payload, headers)

res = conn.getresponse()
data = res.read()

print(data.decode("utf-8"))
```

PHP

```php
<?php

$curl = curl_init();

curl_setopt_array($curl, [
	CURLOPT_URL => "https://www.kdocs.cn/api/v3/ide/file/:file_id/script/:script_id/sync_task",
	CURLOPT_RETURNTRANSFER => true,
	CURLOPT_ENCODING => "",
	CURLOPT_MAXREDIRS => 10,
	CURLOPT_TIMEOUT => 30,
	CURLOPT_HTTP_VERSION => CURL_HTTP_VERSION_1_1,
	CURLOPT_CUSTOMREQUEST => "POST",
	CURLOPT_POSTFIELDS => "{\"Context\":{\"argv\":{},\"sheet_name\":\"表名\",\"range\":\"$B$156\"}}",
	CURLOPT_HTTPHEADER => [
		"AirScript-Token: xxx",
		"Content-Type: application/json"
	],
]);

$response = curl_exec($curl);
$err = curl_error($curl);

curl_close($curl);

if ($err) {
	echo "cURL Error #:" . $err;
} else {
	echo $response;
}
```

JS

```javascript
const data = JSON.stringify({
	"Context": {
		"argv": {},
		"sheet_name": "表名",
		"range": "$B$156"
	}
});

const xhr = new XMLHttpRequest();
xhr.withCredentials = true;

xhr.addEventListener("readystatechange", function () {
	if (this.readyState === this.DONE) {
		console.log(this.responseText);
	}
});

xhr.open("POST", "https://www.kdocs.cn/api/v3/ide/file/:file_id/script/:script_id/sync_task");
xhr.setRequestHeader("Content-Type", "application/json");
xhr.setRequestHeader("AirScript-Token", "xxx");

xhr.send(data);
```

Node.js

```javascript
const http = require("https");

const options = {
	"method": "POST",
	"hostname": "www.kdocs.cn",
	"port": null,
	"path": "/api/v3/ide/file/:file_id/script/:script_id/sync_task",
	"headers": {
		"Content-Type": "application/json",
		"AirScript-Token": "xxx"
	}
};

const req = http.request(options, function (res) {
	const chunks = [];

	res.on("data", function (chunk) {
		chunks.push(chunk);
	});

	res.on("end", function () {
		const body = Buffer.concat(chunks);
		console.log(body.toString());
	});
});

req.write(JSON.stringify({Context: {argv: {}, sheet_name: '表名', range: '$B$156'}}));
req.end();
```

C

```c
CURL *hnd = curl_easy_init();

curl_easy_setopt(hnd, CURLOPT_CUSTOMREQUEST, "POST");
curl_easy_setopt(hnd, CURLOPT_URL, "https://www.kdocs.cn/api/v3/ide/file/:file_id/script/:script_id/sync_task");

struct curl_slist *headers = NULL;
headers = curl_slist_append(headers, "Content-Type: application/json");
headers = curl_slist_append(headers, "AirScript-Token: xxx");
curl_easy_setopt(hnd, CURLOPT_HTTPHEADER, headers);

curl_easy_setopt(hnd, CURLOPT_POSTFIELDS, "{\"Context\":{\"argv\":{},\"sheet_name\":\"表名\",\"range\":\"$B$156\"}}");

CURLcode ret = curl_easy_perform(hnd);
```

C#

```csharp
var client = new RestClient("https://www.kdocs.cn/api/v3/ide/file/:file_id/script/:script_id/sync_task");
var request = new RestRequest(Method.POST);
request.AddHeader("Content-Type", "application/json");
request.AddHeader("AirScript-Token", "xxx");
request.AddParameter("application/json", "{\"Context\":{\"argv\":{},\"sheet_name\":\"表名\",\"range\":\"$B$156\"}}", ParameterType.RequestBody);
IRestResponse response = client.Execute(request);
```

### 返回示例

```json
{
  "data": {
    "logs": [
      {
        "filename": "<system>",
        "timestamp": "16:44:08.271",
        "unix_time": 1690274648271,
        "level": "info",
        "args": ["脚本环境初始化..."]
      },
      {
        "filename": "<system>",
        "timestamp": "16:44:08.953",
        "unix_time": 1690274648953,
        "level": "info",
        "args": ["已开始执行"]
      },
      {
        "filename": "未命名脚本.js:1:9",
        "timestamp": "16:44:08.968",
        "unix_time": 1690274648968,
        "level": "info",
        "args": ["打印参数A：111"]
      },
      {
        "filename": "<system>",
        "timestamp": "16:44:08.969",
        "unix_time": 1690274648969,
        "level": "info",
        "args": ["执行完毕"]
      }
    ],
    "result": "[Undefined]"
  },
  "error": "",
  "status": "finished"
}
```

## 异步执行脚本

`POST /api/v3/ide/file/:file_id/script/:script_id/task`

### Header 参数

| 参数 | 必须 | 类型 | 说明 |
| --- | --- | --- | --- |
| Content-Type | 是 | string | `application/json` |
| AirScript-Token | 是 | string | 传入您通过 AirScript 编辑器生成的**脚本令牌**（APIToken） |

### path 参数

| 参数 | 必须 | 类型 | 说明 |
| --- | --- | --- | --- |
| script_id | 是 | string | 脚本的 ID |
| file_id | 是 | string | 运行脚本的文件 ID |

### body 参数

| 参数 | 必须 | 类型 | 说明 |
| --- | --- | --- | --- |
| Context | 是 | Object | 运行时的上下文参数 |
| Context.argv | 否 | Object | 传入的上下文参数对象，比如传入`{name: 'xiaomeng', age: 18}`，在 AS 代码中可通过`Context.argv.name`获取到传入的值 |
| Context.sheet_name | 否 | string | et,ksheet 运行时所在表名 |
| Context.range | 否 | string | et,ksheet 运行时所在区域，例如`$B$156` |
| Context.link_from | 否 | string | et,ksheet 点击超链接所在单元格 |
| Context.db_active_view | 否 | string | db 运行时所在 view 名 |
| Context.db_selection | 否 | string | db 运行时所在选区 |

### 返回参数

| 参数 | 必须 | 类型 | 说明 |
| --- | --- | --- | --- |
| task_id | 是 | string | 运行的任务 Id，用于轮循运行结果 |
| task_type | 是 | string | 任务类型 |

### 请求示例

Shell

```shell
curl --request POST \
	--url https://www.kdocs.cn/api/v3/ide/file/:file_id/script/:script_id/task \
	--header 'AirScript-Token: xxx' \
	--header 'Content-Type: application/json' \
	--data '{"Context":{"argv":{},"sheet_name":"表名","range":"$B$156"}}'
```

Java

```java
OkHttpClient client = new OkHttpClient();

MediaType mediaType = MediaType.parse("application/json");
RequestBody body = RequestBody.create(mediaType, "{\"Context\":{\"argv\":{},\"sheet_name\":\"表名\",\"range\":\"$B$156\"}}");
Request request = new Request.Builder()
	.url("https://www.kdocs.cn/api/v3/ide/file/:file_id/script/:script_id/task")
	.post(body)
	.addHeader("Content-Type", "application/json")
	.addHeader("AirScript-Token", "xxx")
	.build();

Response response = client.newCall(request).execute();
```

Go

```go
package main

import (
	"fmt"
	"strings"
	"net/http"
	"io/ioutil"
)

func main() {

	url := "https://www.kdocs.cn/api/v3/ide/file/:file_id/script/:script_id/task"

	payload := strings.NewReader("{\"Context\":{\"argv\":{},\"sheet_name\":\"表名\",\"range\":\"$B$156\"}}")

	req, _ := http.NewRequest("POST", url, payload)

	req.Header.Add("Content-Type", "application/json")
	req.Header.Add("AirScript-Token", "xxx")

	res, _ := http.DefaultClient.Do(req)

	defer res.Body.Close()
	body, _ := ioutil.ReadAll(res.Body)

	fmt.Println(res)
	fmt.Println(string(body))

}
```

Python

```python
import http.client

conn = http.client.HTTPSConnection("www.kdocs.cn")

payload = "{\"Context\":{\"argv\":{},\"sheet_name\":\"表名\",\"range\":\"$B$156\"}}"

headers = {
    'Content-Type': "application/json",
    'AirScript-Token': "xxx"
    }

conn.request("POST", "/api/v3/ide/file/:file_id/script/:script_id/task", payload, headers)

res = conn.getresponse()
data = res.read()

print(data.decode("utf-8"))
```

PHP

```php
<?php

$curl = curl_init();

curl_setopt_array($curl, [
	CURLOPT_URL => "https://www.kdocs.cn/api/v3/ide/file/:file_id/script/:script_id/task",
	CURLOPT_RETURNTRANSFER => true,
	CURLOPT_ENCODING => "",
	CURLOPT_MAXREDIRS => 10,
	CURLOPT_TIMEOUT => 30,
	CURLOPT_HTTP_VERSION => CURL_HTTP_VERSION_1_1,
	CURLOPT_CUSTOMREQUEST => "POST",
	CURLOPT_POSTFIELDS => "{\"Context\":{\"argv\":{},\"sheet_name\":\"表名\",\"range\":\"$B$156\"}}",
	CURLOPT_HTTPHEADER => [
		"AirScript-Token: xxx",
		"Content-Type: application/json"
	],
]);

$response = curl_exec($curl);
$err = curl_error($curl);

curl_close($curl);

if ($err) {
	echo "cURL Error #:" . $err;
} else {
	echo $response;
}
```

JS

```javascript
const data = JSON.stringify({
	"Context": {
		"argv": {},
		"sheet_name": "表名",
		"range": "$B$156"
	}
});

const xhr = new XMLHttpRequest();
xhr.withCredentials = true;

xhr.addEventListener("readystatechange", function () {
	if (this.readyState === this.DONE) {
		console.log(this.responseText);
	}
});

xhr.open("POST", "https://www.kdocs.cn/api/v3/ide/file/:file_id/script/:script_id/task");
xhr.setRequestHeader("Content-Type", "application/json");
xhr.setRequestHeader("AirScript-Token", "xxx");

xhr.send(data);
```

Node.js

```javascript
const http = require("https");

const options = {
	"method": "POST",
	"hostname": "www.kdocs.cn",
	"port": null,
	"path": "/api/v3/ide/file/:file_id/script/:script_id/task",
	"headers": {
		"Content-Type": "application/json",
		"AirScript-Token": "xxx"
	}
};

const req = http.request(options, function (res) {
	const chunks = [];

	res.on("data", function (chunk) {
		chunks.push(chunk);
	});

	res.on("end", function () {
		const body = Buffer.concat(chunks);
		console.log(body.toString());
	});
});

req.write(JSON.stringify({Context: {argv: {}, sheet_name: '表名', range: '$B$156'}}));
req.end();
```

C

```c
CURL *hnd = curl_easy_init();

curl_easy_setopt(hnd, CURLOPT_CUSTOMREQUEST, "POST");
curl_easy_setopt(hnd, CURLOPT_URL, "https://www.kdocs.cn/api/v3/ide/file/:file_id/script/:script_id/task");

struct curl_slist *headers = NULL;
headers = curl_slist_append(headers, "Content-Type: application/json");
headers = curl_slist_append(headers, "AirScript-Token: xxx");
curl_easy_setopt(hnd, CURLOPT_HTTPHEADER, headers);

curl_easy_setopt(hnd, CURLOPT_POSTFIELDS, "{\"Context\":{\"argv\":{},\"sheet_name\":\"表名\",\"range\":\"$B$156\"}}");

CURLcode ret = curl_easy_perform(hnd);
```

C#

```csharp
var client = new RestClient("https://www.kdocs.cn/api/v3/ide/file/:file_id/script/:script_id/task");
var request = new RestRequest(Method.POST);
request.AddHeader("Content-Type", "application/json");
request.AddHeader("AirScript-Token", "xxx");
request.AddParameter("application/json", "{\"Context\":{\"argv\":{},\"sheet_name\":\"表名\",\"range\":\"$B$156\"}}", ParameterType.RequestBody);
IRestResponse response = client.Execute(request);
```

### 返回示例

```json
{
  "data": {
    "task_id": "GN/KU3B3BG84MdCjraN5mukx0Rt5Sp1eJ9k2qClmcaOkkF3PUVNDOYPY7Kz4aQMXSvXn9N08QabldRKjPfzii87fuGYydIuK2la2HMfcxmGK1Pf4WcPEflb5xOOkQQEo8fmEbzcobhurYg=="
  },
  "task_id": "GN/KU3B3BG84MdCjraN5mukx0Rt5Sp1eJ9k2qClmcaOkkF3PUVNDOYPY7Kz4aQMXSvXn9N08QabldRKjPfzii87fuGYydIuK2la2HMfcxmGK1Pf4WcPEflb5xOOkQQEo8fmEbzcobhurYg==",
  "task_type": "open_air_script"
}
```

## 获取任务运行情况

`GET /api/v3/script/task`

### query 参数

| 参数 | 必须 | 类型 | 说明 |
| --- | --- | --- | --- |
| task_id | 是 | string | 执行异步任务时返回的 ID |

提示

任务ID为query参数，拼接时请注意先编码下，比如`encodeURIComponent(task_id)`

### 返回参数

| 参数 | 必须 | 类型 | 说明 |
| --- | --- | --- | --- |
| data | 是 | Object | 任务执行数据对象 |
| data.result | 是 | string | 任务执行返回的数据 |
| data.logs | 是 | Array | 任务执行日志 |
| data.logs[i].filename | 是 | string | 执行文件的名称 |
| data.logs[i].timestamp | 是 | string | 执行时间 |
| data.logs[i].unix_time | 是 | number | 执行 unix 时间戳 |
| data.logs[i].level | 是 | string | 日志级别 |
| data.logs[i].args | 是 | string[] | 日志打印参数 |
| status | 是 | string | 任务是否执行完毕 |
| error | 是 | string | 任务执行错误信息 |
| error_details | 否 | object | 错误信息详情对象 |
| error_details.name | 否 | string | 错误信息名称 |
| error_details.msg | 否 | string | 错误信息 |
| error_details.stack | 否 | string[] | 错误信息栈 |
| error_details.unix_time | 否 | number | 错误信息 unix 时间 |

### 请求示例

Shell

```shell
curl --request GET \
	--url https://www.kdocs.cn/api/v3/script/task
```

Java

```java
OkHttpClient client = new OkHttpClient();

Request request = new Request.Builder()
	.url("https://www.kdocs.cn/api/v3/script/task")
	.get()
	.build();

Response response = client.newCall(request).execute();
```

Go

```go
package main

import (
	"fmt"
	"net/http"
	"io/ioutil"
)

func main() {

	url := "https://www.kdocs.cn/api/v3/script/task"

	req, _ := http.NewRequest("GET", url, nil)

	res, _ := http.DefaultClient.Do(req)

	defer res.Body.Close()
	body, _ := ioutil.ReadAll(res.Body)

	fmt.Println(res)
	fmt.Println(string(body))

}
```

Python

```python
import http.client

conn = http.client.HTTPSConnection("www.kdocs.cn")

conn.request("GET", "/api/v3/script/task")

res = conn.getresponse()
data = res.read()

print(data.decode("utf-8"))
```

PHP

```php
<?php

$curl = curl_init();

curl_setopt_array($curl, [
	CURLOPT_URL => "https://www.kdocs.cn/api/v3/script/task",
	CURLOPT_RETURNTRANSFER => true,
	CURLOPT_ENCODING => "",
	CURLOPT_MAXREDIRS => 10,
	CURLOPT_TIMEOUT => 30,
	CURLOPT_HTTP_VERSION => CURL_HTTP_VERSION_1_1,
	CURLOPT_CUSTOMREQUEST => "GET",
]);

$response = curl_exec($curl);
$err = curl_error($curl);

curl_close($curl);

if ($err) {
	echo "cURL Error #:" . $err;
} else {
	echo $response;
}
```

JS

```javascript
const data = null;

const xhr = new XMLHttpRequest();
xhr.withCredentials = true;

xhr.addEventListener("readystatechange", function () {
	if (this.readyState === this.DONE) {
		console.log(this.responseText);
	}
});

xhr.open("GET", "https://www.kdocs.cn/api/v3/script/task");

xhr.send(data);
```

Node.js

```javascript
const http = require("https");

const options = {
	"method": "GET",
	"hostname": "www.kdocs.cn",
	"port": null,
	"path": "/api/v3/script/task",
	"headers": {}
};

const req = http.request(options, function (res) {
	const chunks = [];

	res.on("data", function (chunk) {
		chunks.push(chunk);
	});

	res.on("end", function () {
		const body = Buffer.concat(chunks);
		console.log(body.toString());
	});
});

req.end();
```

C

```c
CURL *hnd = curl_easy_init();

curl_easy_setopt(hnd, CURLOPT_CUSTOMREQUEST, "GET");
curl_easy_setopt(hnd, CURLOPT_URL, "https://www.kdocs.cn/api/v3/script/task");

CURLcode ret = curl_easy_perform(hnd);
```

C#

```csharp
var client = new RestClient("https://www.kdocs.cn/api/v3/script/task");
var request = new RestRequest(Method.GET);
IRestResponse response = client.Execute(request);
```

### 返回示例

```json
{
  "data": {
    "logs": [
      {
        "filename": "<system>",
        "timestamp": "17:05:16.164",
        "unix_time": 1692090316164,
        "level": "info",
        "args": ["脚本环境初始化..."]
      }
    ],
    "result": null
  },
  "error": "Unexpected token (1:91)",
  "error_details": {
    "name": "SyntaxError",
    "msg": "Unexpected token (1:91)",
    "stack": ["    at 未命名脚本.js:1:91"],
    "unix_time": 1692090318372
  },
  "status": "finished"
}
```
