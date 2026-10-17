---
title: "API Documentation"
source: https://airsheet.wps.cn/docs/en-US/apitoken/api.html
updated: 1789700916000
hash: DvjidoAK
---

# API Documentation

After successfully generating a script token, you can execute scripts through HTTP interface. We provide two script execution interfaces for developers: **Synchronous Execution** and **Asynchronous Execution**.

Comparatively, the former is simpler to use. The interface returns the execution result directly after calling, suitable for scenarios with normal execution time. The latter is slightly more complex. The interface does not return the final execution result after calling, but immediately returns a `task_id`. You need to poll for the script execution log based on this `task_id` without waiting synchronously for the result to block the business process. This interface is suitable for scenarios with longer execution time.

Regardless of which interface you use, you must first obtain the file ID and script ID. Please enter the script editor first and copy the webhook link from the more menu in the sidebar list.

## Synchronous Script Execution

`POST /api/v3/ide/file/:file_id/script/:script_id/sync_task`

### Header Parameters

| Parameter | Required | Type | Description |
| --- | --- | --- | --- |
| Content-Type | Yes | string | `application/json` |
| AirScript-Token | Yes | string | Pass in your **Script Token** (APIToken) generated via AirScript editor |

### Path Parameters

| Parameter | Required | Type | Description |
| --- | --- | --- | --- |
| script_id | Yes | string | Script ID |
| file_id | Yes | string | File ID where script runs |

### Body Parameters

| Parameter | Required | Type | Description |
| --- | --- | --- | --- |
| Context | Yes | Object | Runtime context parameters |
| Context.argv | No | Object | Context parameter object passed in, e.g., pass `{name: 'John', age: 18}`, accessible in AS code via `Context.argv.name` |
| Context.sheet_name | No | string | Sheet name at runtime for et,ksheet |
| Context.range | No | string | Range at runtime for et,ksheet, e.g., `$B$156` |
| Context.link_from | No | string | Cell where hyperlink was clicked for et,ksheet |
| Context.db_active_view | No | string | View name at runtime for db |
| Context.db_selection | No | string | Selection at runtime for db |

### Response Parameters

| Parameter | Required | Type | Description |
| --- | --- | --- | --- |
| data | Yes | Object | Task execution data object |
| data.result | Yes | string | Data returned by task |
| data.logs | Yes | Array | Task execution logs |
| data.logs[i].filename | Yes | string | Executed file name |
| data.logs[i].timestamp | Yes | string | Execution time |
| data.logs[i].unix_time | Yes | number | Unix timestamp |
| data.logs[i].level | Yes | string | Log level |
| data.logs[i].args | Yes | string[] | Log print parameters |
| status | Yes | string | Whether task is completed |
| error | Yes | string | Task execution error message |
| error_details | No | object | Error details object |
| error_details.name | No | string | Error name |
| error_details.msg | No | string | Error message |
| error_details.stack | No | string[] | Error stack |
| error_details.unix_time | No | number | Error unix time |

### Request Example

Shell

```shell
curl --request POST \
	--url https://www.wps.com/api/v3/ide/file/:file_id/script/:script_id/sync_task \
	--header 'AirScript-Token: xxx' \
	--header 'Content-Type: application/json' \
	--data '{"Context":{"argv":{},"sheet_name":"SheetName","range":"$B$156"}}'
```

Java

```java
OkHttpClient client = new OkHttpClient();

MediaType mediaType = MediaType.parse("application/json");
RequestBody body = RequestBody.create(mediaType, "{\"Context\":{\"argv\":{},\"sheet_name\":\"SheetName\",\"range\":\"$B$156\"}}");
Request request = new Request.Builder()
	.url("https://www.wps.com/api/v3/ide/file/:file_id/script/:script_id/sync_task")
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

	url := "https://www.wps.com/api/v3/ide/file/:file_id/script/:script_id/sync_task"

	payload := strings.NewReader("{\"Context\":{\"argv\":{},\"sheet_name\":\"SheetName\",\"range\":\"$B$156\"}}")

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

conn = http.client.HTTPSConnection("www.wps.com")

payload = "{\"Context\":{\"argv\":{},\"sheet_name\":\"SheetName\",\"range\":\"$B$156\"}}"

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
	CURLOPT_URL => "https://www.wps.com/api/v3/ide/file/:file_id/script/:script_id/sync_task",
	CURLOPT_RETURNTRANSFER => true,
	CURLOPT_ENCODING => "",
	CURLOPT_MAXREDIRS => 10,
	CURLOPT_TIMEOUT => 30,
	CURLOPT_HTTP_VERSION => CURL_HTTP_VERSION_1_1,
	CURLOPT_CUSTOMREQUEST => "POST",
	CURLOPT_POSTFIELDS => "{\"Context\":{\"argv\":{},\"sheet_name\":\"SheetName\",\"range\":\"$B$156\"}}",
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
		"sheet_name": "SheetName",
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

xhr.open("POST", "https://www.wps.com/api/v3/ide/file/:file_id/script/:script_id/sync_task");
xhr.setRequestHeader("Content-Type", "application/json");
xhr.setRequestHeader("AirScript-Token", "xxx");

xhr.send(data);
```

Node.js

```javascript
const http = require("https");

const options = {
	"method": "POST",
	"hostname": "www.wps.com",
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

req.write(JSON.stringify({Context: {argv: {}, sheet_name: 'SheetName', range: '$B$156'}}));
req.end();
```

C

```c
CURL *hnd = curl_easy_init();

curl_easy_setopt(hnd, CURLOPT_CUSTOMREQUEST, "POST");
curl_easy_setopt(hnd, CURLOPT_URL, "https://www.wps.com/api/v3/ide/file/:file_id/script/:script_id/sync_task");

struct curl_slist *headers = NULL;
headers = curl_slist_append(headers, "Content-Type: application/json");
headers = curl_slist_append(headers, "AirScript-Token: xxx");
curl_easy_setopt(hnd, CURLOPT_HTTPHEADER, headers);

curl_easy_setopt(hnd, CURLOPT_POSTFIELDS, "{\"Context\":{\"argv\":{},\"sheet_name\":\"SheetName\",\"range\":\"$B$156\"}}");

CURLcode ret = curl_easy_perform(hnd);
```

C#

```csharp
var client = new RestClient("https://www.wps.com/api/v3/ide/file/:file_id/script/:script_id/sync_task");
var request = new RestRequest(Method.POST);
request.AddHeader("Content-Type", "application/json");
request.AddHeader("AirScript-Token", "xxx");
request.AddParameter("application/json", "{\"Context\":{\"argv\":{},\"sheet_name\":\"SheetName\",\"range\":\"$B$156\"}}", ParameterType.RequestBody);
IRestResponse response = client.Execute(request);
```

### Response Example

```json
{
  "data": {
    "logs": [
      {
        "filename": "<system>",
        "timestamp": "16:44:08.271",
        "unix_time": 1690274648271,
        "level": "info",
        "args": ["Script environment initializing..."]
      },
      {
        "filename": "<system>",
        "timestamp": "16:44:08.953",
        "unix_time": 1690274648953,
        "level": "info",
        "args": ["Execution started"]
      },
      {
        "filename": "UntitledScript.js:1:9",
        "timestamp": "16:44:08.968",
        "unix_time": 1690274648968,
        "level": "info",
        "args": ["Print parameter A: 111"]
      },
      {
        "filename": "<system>",
        "timestamp": "16:44:08.969",
        "unix_time": 1690274648969,
        "level": "info",
        "args": ["Execution completed"]
      }
    ],
    "result": "[Undefined]"
  },
  "error": "",
  "status": "finished"
}
```

## Asynchronous Script Execution

`POST /api/v3/ide/file/:file_id/script/:script_id/task`

### Header Parameters

| Parameter | Required | Type | Description |
| --- | --- | --- | --- |
| Content-Type | Yes | string | `application/json` |
| AirScript-Token | Yes | string | Pass in your **Script Token** (APIToken) generated via AirScript editor |

### Path Parameters

| Parameter | Required | Type | Description |
| --- | --- | --- | --- |
| script_id | Yes | string | Script ID |
| file_id | Yes | string | File ID where script runs |

### Body Parameters

| Parameter | Required | Type | Description |
| --- | --- | --- | --- |
| Context | Yes | Object | Runtime context parameters |
| Context.argv | No | Object | Context parameter object passed in, e.g., pass `{name: 'John', age: 18}`, accessible in AS code via `Context.argv.name` |
| Context.sheet_name | No | string | Sheet name at runtime for et,ksheet |
| Context.range | No | string | Range at runtime for et,ksheet, e.g., `$B$156` |
| Context.link_from | No | string | Cell where hyperlink was clicked for et,ksheet |
| Context.db_active_view | No | string | View name at runtime for db |
| Context.db_selection | No | string | Selection at runtime for db |

### Response Parameters

| Parameter | Required | Type | Description |
| --- | --- | --- | --- |
| task_id | Yes | string | Running task ID for polling results |
| task_type | Yes | string | Task type |

### Request Example

Shell

```shell
curl --request POST \
	--url https://www.wps.com/api/v3/ide/file/:file_id/script/:script_id/task \
	--header 'AirScript-Token: xxx' \
	--header 'Content-Type: application/json' \
	--data '{"Context":{"argv":{},"sheet_name":"SheetName","range":"$B$156"}}'
```

Java

```java
OkHttpClient client = new OkHttpClient();

MediaType mediaType = MediaType.parse("application/json");
RequestBody body = RequestBody.create(mediaType, "{\"Context\":{\"argv\":{},\"sheet_name\":\"SheetName\",\"range\":\"$B$156\"}}");
Request request = new Request.Builder()
	.url("https://www.wps.com/api/v3/ide/file/:file_id/script/:script_id/task")
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

	url := "https://www.wps.com/api/v3/ide/file/:file_id/script/:script_id/task"

	payload := strings.NewReader("{\"Context\":{\"argv\":{},\"sheet_name\":\"SheetName\",\"range\":\"$B$156\"}}")

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

conn = http.client.HTTPSConnection("www.wps.com")

payload = "{\"Context\":{\"argv\":{},\"sheet_name\":\"SheetName\",\"range\":\"$B$156\"}}"

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
	CURLOPT_URL => "https://www.wps.com/api/v3/ide/file/:file_id/script/:script_id/task",
	CURLOPT_RETURNTRANSFER => true,
	CURLOPT_ENCODING => "",
	CURLOPT_MAXREDIRS => 10,
	CURLOPT_TIMEOUT => 30,
	CURLOPT_HTTP_VERSION => CURL_HTTP_VERSION_1_1,
	CURLOPT_CUSTOMREQUEST => "POST",
	CURLOPT_POSTFIELDS => "{\"Context\":{\"argv\":{},\"sheet_name\":\"SheetName\",\"range\":\"$B$156\"}}",
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
		"sheet_name": "SheetName",
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

xhr.open("POST", "https://www.wps.com/api/v3/ide/file/:file_id/script/:script_id/task");
xhr.setRequestHeader("Content-Type", "application/json");
xhr.setRequestHeader("AirScript-Token", "xxx");

xhr.send(data);
```

Node.js

```javascript
const http = require("https");

const options = {
	"method": "POST",
	"hostname": "www.wps.com",
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

req.write(JSON.stringify({Context: {argv: {}, sheet_name: 'SheetName', range: '$B$156'}}));
req.end();
```

C

```c
CURL *hnd = curl_easy_init();

curl_easy_setopt(hnd, CURLOPT_CUSTOMREQUEST, "POST");
curl_easy_setopt(hnd, CURLOPT_URL, "https://www.wps.com/api/v3/ide/file/:file_id/script/:script_id/task");

struct curl_slist *headers = NULL;
headers = curl_slist_append(headers, "Content-Type: application/json");
headers = curl_slist_append(headers, "AirScript-Token: xxx");
curl_easy_setopt(hnd, CURLOPT_HTTPHEADER, headers);

curl_easy_setopt(hnd, CURLOPT_POSTFIELDS, "{\"Context\":{\"argv\":{},\"sheet_name\":\"SheetName\",\"range\":\"$B$156\"}}");

CURLcode ret = curl_easy_perform(hnd);
```

C#

```csharp
var client = new RestClient("https://www.wps.com/api/v3/ide/file/:file_id/script/:script_id/task");
var request = new RestRequest(Method.POST);
request.AddHeader("Content-Type", "application/json");
request.AddHeader("AirScript-Token", "xxx");
request.AddParameter("application/json", "{\"Context\":{\"argv\":{},\"sheet_name\":\"SheetName\",\"range\":\"$B$156\"}}", ParameterType.RequestBody);
IRestResponse response = client.Execute(request);
```

### Response Example

```json
{
  "data": {
    "task_id": "GN/KU3B3BG84MdCjraN5mukx0Rt5Sp1eJ9k2qClmcaOkkF3PUVNDOYPY7Kz4aQMXSvXn9N08QabldRKjPfzii87fuGYydIuK2la2HMfcxmGK1Pf4WcPEflb5xOOkQQEo8fmEbzcobhurYg=="
  },
  "task_id": "GN/KU3B3BG84MdCjraN5mukx0Rt5Sp1eJ9k2qClmcaOkkF3PUVNDOYPY7Kz4aQMXSvXn9N08QabldRKjPfzii87fuGYydIuK2la2HMfcxmGK1Pf4WcPEflb5xOOkQQEo8fmEbzcobhurYg==",
  "task_type": "open_air_script"
}
```

## Get Task Status

`GET /api/v3/script/task`

### Query Parameters

| Parameter | Required | Type | Description |
| --- | --- | --- | --- |
| task_id | Yes | string | ID returned when executing async task |

Tip

Task ID is a query parameter. Please encode it when concatenating, e.g., `encodeURIComponent(task_id)`

### Response Parameters

| Parameter | Required | Type | Description |
| --- | --- | --- | --- |
| data | Yes | Object | Task execution data object |
| data.result | Yes | string | Data returned by task |
| data.logs | Yes | Array | Task execution logs |
| data.logs[i].filename | Yes | string | Executed file name |
| data.logs[i].timestamp | Yes | string | Execution time |
| data.logs[i].unix_time | Yes | number | Unix timestamp |
| data.logs[i].level | Yes | string | Log level |
| data.logs[i].args | Yes | string[] | Log print parameters |
| status | Yes | string | Whether task is completed |
| error | Yes | string | Task execution error message |
| error_details | No | object | Error details object |
| error_details.name | No | string | Error name |
| error_details.msg | No | string | Error message |
| error_details.stack | No | string[] | Error stack |
| error_details.unix_time | No | number | Error unix time |

### Request Example

Shell

```shell
curl --request GET \
	--url https://www.wps.com/api/v3/script/task
```

Java

```java
OkHttpClient client = new OkHttpClient();

Request request = new Request.Builder()
	.url("https://www.wps.com/api/v3/script/task")
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

	url := "https://www.wps.com/api/v3/script/task"

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

conn = http.client.HTTPSConnection("www.wps.com")

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
	CURLOPT_URL => "https://www.wps.com/api/v3/script/task",
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

xhr.open("GET", "https://www.wps.com/api/v3/script/task");

xhr.send(data);
```

Node.js

```javascript
const http = require("https");

const options = {
	"method": "GET",
	"hostname": "www.wps.com",
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
curl_easy_setopt(hnd, CURLOPT_URL, "https://www.wps.com/api/v3/script/task");

CURLcode ret = curl_easy_perform(hnd);
```

C#

```csharp
var client = new RestClient("https://www.wps.com/api/v3/script/task");
var request = new RestRequest(Method.GET);
IRestResponse response = client.Execute(request);
```

### Response Example

```json
{
  "data": {
    "logs": [
      {
        "filename": "<system>",
        "timestamp": "17:05:16.164",
        "unix_time": 1692090316164,
        "level": "info",
        "args": ["Script environment initializing..."]
      }
    ],
    "result": null
  },
  "error": "Unexpected token (1:91)",
  "error_details": {
    "name": "SyntaxError",
    "msg": "Unexpected token (1:91)",
    "stack": ["    at UntitledScript.js:1:91"],
    "unix_time": 1692090318372
  },
  "status": "finished"
}
```
