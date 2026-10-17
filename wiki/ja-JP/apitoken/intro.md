---
title: "スクリプトトークン（APIToken）"
source: https://airsheet.wps.cn/docs/ja-JP/apitoken/intro.html
updated: 1789700916000
hash: CTzmaR5H
---

## スクリプトトークンの作成方法

1. スマートスプレッドシートに入り、スクリプトエディターを開き、ツールバーの［スクリプトトークン（ベータ）］ボタンをクリックします。
2. まだスクリプトトークンを作成したことがない場合は、作成に関する注意が表示されます。［理解しました］にチェックを入れ、［スクリプトトークンを作成］をクリックします。
3. 一定のセキュリティのため、本人確認が完了していない場合は、先に本人確認の手続きを行う必要があります。
4. 作成に成功すると、個人用のスクリプトトークンが表示されます。
5. 確認が完了すると、固有の個人用トークンがすぐに発行されます。**すぐにコピーして安全な場所に保管してください**。第三者と共有しないでください。ドキュメントへのアクセスに利用できるためです。

## スクリプトトークンの使用方法

スクリプトトークンは、外部からスクリプトを実行するための資格情報です。自分のスクリプトトークンを発行したら、そのトークンを使ってスクリプトの実行・呼び出しを始められます。

まずスクリプトエディターを開き、サイドバー内の任意のドキュメントスクリプトのその他メニューから、スクリプトの webhook リンクをコピーします。

注意

**重要な制限：** 現時点では、webhook リンクは「ドキュメント共有スクリプト」に保存されているスクリプトでのみ利用できます。「マイスクリプト」に保存されているスクリプト向けの webhook は、今後のアップデートで対応予定です。また、その他メニューには、スクリプトの名前変更、削除、スプレッドシートのセルへの直接挿入などの便利なオプションも含まれています。

コピーされるリンクの内容は次のとおりです。

`https://www.wps.com/api/v3/ide/file/caEkI6K5RDG2/script/V2-5rSiBiN7y5xdOd5x5ZYI2r/sync_task`

リンクには、対象スクリプトのスクリプト ID とドキュメントのファイル ID が含まれています。この URL にリクエストを送ることで、対応するドキュメントの読み取り・編集ができます。呼び出す際は、リクエストヘッダーに `AirScript-Token` を設定し、値として自分のスクリプトトークンを指定する必要があります。詳しくは [APIドキュメント](api.md) を参照してください。

ここでは、対象スクリプトのコードが次のものであると仮定します。セル A1 の値を `AirScript` に変更し、オブジェクトを返します。

```js
Application.Range('A1').Value2 = 'AirScript'
return {
  name: 'ジン・シャオメン',
  age: 17
}
```

スクリプトトークンと webhook を使って、次のような HTTP リクエストを組み立てます。

Shell

```shell
curl --request POST \
	--url https://www.wps.com/api/v3/ide/file/caEkI6K5RDG2/script/V2-5rSiBiN7y5xdOd5x5ZYI2r/sync_task \
	--header 'AirScript-Token: xxx' \
	--header 'Content-Type: application/json' \
	--data '{"Context":{"argv":{}}}'
```

Java

```java
OkHttpClient client = new OkHttpClient();

MediaType mediaType = MediaType.parse("application/json");
RequestBody body = RequestBody.create(mediaType, "{\"Context\":{\"argv\":{}}}");
Request request = new Request.Builder()
	.url("https://www.wps.com/api/v3/ide/file/caEkI6K5RDG2/script/V2-5rSiBiN7y5xdOd5x5ZYI2r/sync_task")
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

	url := "https://www.wps.com/api/v3/ide/file/caEkI6K5RDG2/script/V2-5rSiBiN7y5xdOd5x5ZYI2r/sync_task"

	payload := strings.NewReader("{\"Context\":{\"argv\":{}}}")

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

payload = "{\"Context\":{\"argv\":{}}}"

headers = {
    'Content-Type': "application/json",
    'AirScript-Token': "xxx"
    }

conn.request("POST", "/api/v3/ide/file/caEkI6K5RDG2/script/V2-5rSiBiN7y5xdOd5x5ZYI2r/sync_task", payload, headers)

res = conn.getresponse()
data = res.read()

print(data.decode("utf-8"))
```

PHP

```php
<?php

$curl = curl_init();

curl_setopt_array($curl, [
	CURLOPT_URL => "https://www.wps.com/api/v3/ide/file/caEkI6K5RDG2/script/V2-5rSiBiN7y5xdOd5x5ZYI2r/sync_task",
	CURLOPT_RETURNTRANSFER => true,
	CURLOPT_ENCODING => "",
	CURLOPT_MAXREDIRS => 10,
	CURLOPT_TIMEOUT => 30,
	CURLOPT_HTTP_VERSION => CURL_HTTP_VERSION_1_1,
	CURLOPT_CUSTOMREQUEST => "POST",
	CURLOPT_POSTFIELDS => "{\"Context\":{\"argv\":{}}}",
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
		"argv": {}
	}
});

const xhr = new XMLHttpRequest();
xhr.withCredentials = true;

xhr.addEventListener("readystatechange", function () {
	if (this.readyState === this.DONE) {
		console.log(this.responseText);
	}
});

xhr.open("POST", "https://www.wps.com/api/v3/ide/file/caEkI6K5RDG2/script/V2-5rSiBiN7y5xdOd5x5ZYI2r/sync_task");
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
	"path": "/api/v3/ide/file/caEkI6K5RDG2/script/V2-5rSiBiN7y5xdOd5x5ZYI2r/sync_task",
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

req.write(JSON.stringify({Context: {argv: {}}}));
req.end();
```

C

```c
CURL *hnd = curl_easy_init();

curl_easy_setopt(hnd, CURLOPT_CUSTOMREQUEST, "POST");
curl_easy_setopt(hnd, CURLOPT_URL, "https://www.wps.com/api/v3/ide/file/caEkI6K5RDG2/script/V2-5rSiBiN7y5xdOd5x5ZYI2r/sync_task");

struct curl_slist *headers = NULL;
headers = curl_slist_append(headers, "Content-Type: application/json");
headers = curl_slist_append(headers, "AirScript-Token: xxx");
curl_easy_setopt(hnd, CURLOPT_HTTPHEADER, headers);

curl_easy_setopt(hnd, CURLOPT_POSTFIELDS, "{\"Context\":{\"argv\":{}}}");

CURLcode ret = curl_easy_perform(hnd);
```

C#

```csharp
var client = new RestClient("https://www.wps.com/api/v3/ide/file/caEkI6K5RDG2/script/V2-5rSiBiN7y5xdOd5x5ZYI2r/sync_task");
var request = new RestRequest(Method.POST);
request.AddHeader("Content-Type", "application/json");
request.AddHeader("AirScript-Token", "xxx");
request.AddParameter("application/json", "{\"Context\":{\"argv\":{}}}", ParameterType.RequestBody);
IRestResponse response = client.Execute(request);
```

リクエストが成功すると、次のようなデータが返されます。内容には主にスクリプト実行のログ情報と、コードで返したデータが含まれます。スクリプトのコードに誤りがある場合は、ログに対応するエラー情報も出力されます。

```json
{
  "data": {
    "logs": [
      {
        "filename": "<system>",
        "timestamp": "12:03:20.711",
        "unix_time": 1691726600711,
        "level": "info",
        "args": ["スクリプト環境を初期化しています..."]
      },
      {
        "filename": "<system>",
        "timestamp": "12:03:22.129",
        "unix_time": 1691726602129,
        "level": "info",
        "args": ["実行を開始しました"]
      },
      {
        "filename": "<system>",
        "timestamp": "12:03:22.312",
        "unix_time": 1691726602312,
        "level": "info",
        "args": ["実行が完了しました"]
      }
    ],
    "result": {
      "age": 17,
      "name": "ジン・シャオメン"
    }
  },
  "error": "",
  "status": "finished"
}
```

## 注意事項

1. スクリプトトークンは第三者がプラットフォーム上のサーバーリソースにアクセスできるようにするため、**セキュリティ**向上のため、本人確認の完了を求めています（既に確認済みの場合は省略できます）。
2. スクリプトトークンは外部スクリプト実行の資格情報であり、**個人の機密情報**にあたります。スクリプトトークンとスクリプトの webhook を組み合わせると、対応するファイルの読み取り・編集が可能になります。適切に管理し、外部に共有しないでください。
3. スクリプトトークンはユーザーに紐づきます。1 ユーザーあたり**最大 1 つ**までです。新しいトークンを作成する場合は、先に古いトークンを削除する必要があります（新しく作成したスクリプトトークンは、元のトークンと異なるものである必要があります）。
4. スクリプトトークンの有効期限は既定で 180 日です。ユーザーは作成時に手動で延長でき、回数の制限なく複数回延長できます。
