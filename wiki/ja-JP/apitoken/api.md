---
title: "APIドキュメント"
source: https://airsheet.wps.cn/docs/ja-JP/apitoken/api.html
updated: 1789700916000
hash: D-lDVkX2
---

# APIドキュメント

スクリプトトークンを発行したら、HTTP インターフェース経由でスクリプトを実行できます。開発者向けに **同期実行** と **非同期実行** の 2 種類のスクリプト実行インターフェースを提供しています。

比較すると、前者は扱いが簡単で、呼び出し後すぐに実行結果が返り、実行時間が通常程度のシーンに適しています。後者はやや複雑で、呼び出し後すぐに最終的な実行結果は返らず、即座に `task_id` が返されます。この `task_id` に基づいてスクリプト実行ログをポーリングし、結果を同期的に待って業務フローをブロックする必要はありません。実行時間が長いシーン向けのインターフェースです。

どちらのインターフェースを使う場合も、先にファイル ID とスクリプト ID を取得する必要があります。スクリプトエディターを開き、サイドバーの一覧でその他メニューから webhook リンクをコピーしてください。

## 同期スクリプト実行

`POST /api/v3/ide/file/:file_id/script/:script_id/sync_task`

### ヘッダーパラメーター

| パラメーター | 必須 | 型 | 説明 |
| --- | --- | --- | --- |
| Content-Type | はい | string | `application/json` |
| AirScript-Token | はい | string | AirScript エディターで発行した **スクリプトトークン**（APIToken）を渡す |

### パスパラメーター

| パラメーター | 必須 | 型 | 説明 |
| --- | --- | --- | --- |
| script_id | はい | string | スクリプト ID |
| file_id | はい | string | スクリプトが実行されるファイル ID |

### ボディパラメーター

| パラメーター | 必須 | 型 | 説明 |
| --- | --- | --- | --- |
| Context | はい | Object | 実行時コンテキストパラメーター |
| Context.argv | いいえ | Object | 渡す Context のパラメーターオブジェクト。例：`{name: 'John', age: 18}`。AS コードでは `Context.argv.name` で参照可能 |
| Context.sheet_name | いいえ | string | et、ksheet の実行時シート名 |
| Context.range | いいえ | string | et、ksheet の実行時 Range。例：`$B$156` |
| Context.link_from | いいえ | string | et、ksheet でハイパーリンクがクリックされたセル |
| Context.db_active_view | いいえ | string | db の実行時ビュー名 |
| Context.db_selection | いいえ | string | db の実行時選択範囲 |

### レスポンスパラメーター

| パラメーター | 必須 | 型 | 説明 |
| --- | --- | --- | --- |
| data | はい | Object | タスク実行データオブジェクト |
| data.result | はい | string | タスクが返したデータ |
| data.logs | はい | Array | タスク実行ログ |
| data.logs[i].filename | はい | string | 実行されたファイル名 |
| data.logs[i].timestamp | はい | string | 実行時刻 |
| data.logs[i].unix_time | はい | number | Unix タイムスタンプ |
| data.logs[i].level | はい | string | ログレベル |
| data.logs[i].args | はい | string[] | ログ出力パラメーター |
| status | はい | string | タスク完了の有無 |
| error | はい | string | タスク実行エラーメッセージ |
| error_details | いいえ | object | エラー詳細オブジェクト |
| error_details.name | いいえ | string | エラー名 |
| error_details.msg | いいえ | string | エラーメッセージ |
| error_details.stack | いいえ | string[] | エラースタック |
| error_details.unix_time | いいえ | number | エラー発生時の unix 時刻 |

### リクエスト例

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

### レスポンス例

```json
{
  "data": {
    "logs": [
      {
        "filename": "<system>",
        "timestamp": "16:44:08.271",
        "unix_time": 1690274648271,
        "level": "info",
        "args": ["スクリプト環境を初期化しています..."]
      },
      {
        "filename": "<system>",
        "timestamp": "16:44:08.953",
        "unix_time": 1690274648953,
        "level": "info",
        "args": ["実行を開始しました"]
      },
      {
        "filename": "UntitledScript.js:1:9",
        "timestamp": "16:44:08.968",
        "unix_time": 1690274648968,
        "level": "info",
        "args": ["パラメーターAを出力: 111"]
      },
      {
        "filename": "<system>",
        "timestamp": "16:44:08.969",
        "unix_time": 1690274648969,
        "level": "info",
        "args": ["実行が完了しました"]
      }
    ],
    "result": "[Undefined]"
  },
  "error": "",
  "status": "finished"
}
```

## 非同期スクリプト実行

`POST /api/v3/ide/file/:file_id/script/:script_id/task`

### ヘッダーパラメーター

| パラメーター | 必須 | 型 | 説明 |
| --- | --- | --- | --- |
| Content-Type | はい | string | `application/json` |
| AirScript-Token | はい | string | AirScript エディターで発行した **スクリプトトークン**（APIToken）を渡す |

### パスパラメーター

| パラメーター | 必須 | 型 | 説明 |
| --- | --- | --- | --- |
| script_id | はい | string | スクリプト ID |
| file_id | はい | string | スクリプトが実行されるファイル ID |

### ボディパラメーター

| パラメーター | 必須 | 型 | 説明 |
| --- | --- | --- | --- |
| Context | はい | Object | 実行時コンテキストパラメーター |
| Context.argv | いいえ | Object | 渡す Context のパラメーターオブジェクト。例：`{name: 'John', age: 18}`。AS コードでは `Context.argv.name` で参照可能 |
| Context.sheet_name | いいえ | string | et、ksheet の実行時シート名 |
| Context.range | いいえ | string | et、ksheet の実行時 Range。例：`$B$156` |
| Context.link_from | いいえ | string | et、ksheet でハイパーリンクがクリックされたセル |
| Context.db_active_view | いいえ | string | db の実行時ビュー名 |
| Context.db_selection | いいえ | string | db の実行時選択範囲 |

### レスポンスパラメーター

| パラメーター | 必須 | 型 | 説明 |
| --- | --- | --- | --- |
| task_id | はい | string | 結果ポーリング用の実行中タスク ID |
| task_type | はい | string | タスク種別 |

### リクエスト例

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

### レスポンス例

```json
{
  "data": {
    "task_id": "GN/KU3B3BG84MdCjraN5mukx0Rt5Sp1eJ9k2qClmcaOkkF3PUVNDOYPY7Kz4aQMXSvXn9N08QabldRKjPfzii87fuGYydIuK2la2HMfcxmGK1Pf4WcPEflb5xOOkQQEo8fmEbzcobhurYg=="
  },
  "task_id": "GN/KU3B3BG84MdCjraN5mukx0Rt5Sp1eJ9k2qClmcaOkkF3PUVNDOYPY7Kz4aQMXSvXn9N08QabldRKjPfzii87fuGYydIuK2la2HMfcxmGK1Pf4WcPEflb5xOOkQQEo8fmEbzcobhurYg==",
  "task_type": "open_air_script"
}
```

## タスクステータスの取得

`GET /api/v3/script/task`

### クエリパラメーター

| パラメーター | 必須 | 型 | 説明 |
| --- | --- | --- | --- |
| task_id | はい | string | 非同期タスク実行時に返された ID |

ヒント

タスク ID はクエリパラメーターです。連結するときはエンコードしてください。例：`encodeURIComponent(task_id)`

### レスポンスパラメーター

| パラメーター | 必須 | 型 | 説明 |
| --- | --- | --- | --- |
| data | はい | Object | タスク実行データオブジェクト |
| data.result | はい | string | タスクが返したデータ |
| data.logs | はい | Array | タスク実行ログ |
| data.logs[i].filename | はい | string | 実行されたファイル名 |
| data.logs[i].timestamp | はい | string | 実行時刻 |
| data.logs[i].unix_time | はい | number | Unix タイムスタンプ |
| data.logs[i].level | はい | string | ログレベル |
| data.logs[i].args | はい | string[] | ログ出力パラメーター |
| status | はい | string | タスク完了の有無 |
| error | はい | string | タスク実行エラーメッセージ |
| error_details | いいえ | object | エラー詳細オブジェクト |
| error_details.name | いいえ | string | エラー名 |
| error_details.msg | いいえ | string | エラーメッセージ |
| error_details.stack | いいえ | string[] | エラースタック |
| error_details.unix_time | いいえ | number | エラー発生時の unix 時刻 |

### リクエスト例

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

### レスポンス例

```json
{
  "data": {
    "logs": [
      {
        "filename": "<system>",
        "timestamp": "17:05:16.164",
        "unix_time": 1692090316164,
        "level": "info",
        "args": ["スクリプト環境を初期化しています..."]
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
