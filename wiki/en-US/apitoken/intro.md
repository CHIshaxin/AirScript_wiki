---
title: "Script Token (APIToken)"
source: https://airsheet.wps.cn/docs/en-US/apitoken/intro.html
updated: 1789700916000
hash: B0iEZifa
---

## How to Create a Script Token?

1. After entering Smart Spreadsheet, open the script editor and click the [Script Token (beta)] button in the toolbar
2. If you have not created a script token before, you will be prompted with notes about creating a script token. Check [I understand], then click [Create Script Token]
3. To ensure a certain level of security, if you have not completed identity verification, you need to complete the identity verification process first
4. After successful creation, you will get your personal script token.
5. Once verified, your unique personal token will be generated immediately. **Copy it right away and store it in a secure place** -- never share it with anyone, as it can be used to access your documents.

## How to Use a Script Token?

Script Token is the credential for external script execution. After successfully generating your own script token, you can start using the script token to execute and call scripts.

First, open the script editor and copy the script's webhook link from the more menu of any document script in the sidebar.

Note

**Important limitation:** Currently, webhook links are only available for scripts saved under "Document Shared Scripts". Webhook support for personal scripts under "My Scripts" is coming in a future update. The menu also includes other useful options like renaming the script, deleting it, or inserting it directly into a spreadsheet cell.

The copied link content is as follows:

`https://www.wps.com/api/v3/ide/file/caEkI6K5RDG2/script/V2-5rSiBiN7y5xdOd5x5ZYI2r/sync_task`

The link already contains the script ID of the current script and the file ID of the document. You can then request this link to read and edit your corresponding documents. Note that when calling, you must set the request header `AirScript-Token` with your script token as the value. For more detailed instructions, please refer to [API Documentation](api.md).

Here we assume the target script code is as follows, modifying the value of cell A1 to `AirScript` and returning an object:

```js
Application.Range('A1').Value2 = 'AirScript'
return {
  name: 'Jin Xiaomeng',
  age: 17
}
```

Through the script token and webhook, we construct an HTTP request as follows:

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

If the request is successful, the following data will be returned. The content mainly includes the log information of script execution and the data returned in the code. If your script code is incorrect, the corresponding error information will also be reflected in the log.

```json
{
  "data": {
    "logs": [
      {
        "filename": "<system>",
        "timestamp": "12:03:20.711",
        "unix_time": 1691726600711,
        "level": "info",
        "args": ["Script environment initializing..."]
      },
      {
        "filename": "<system>",
        "timestamp": "12:03:22.129",
        "unix_time": 1691726602129,
        "level": "info",
        "args": ["Execution started"]
      },
      {
        "filename": "<system>",
        "timestamp": "12:03:22.312",
        "unix_time": 1691726602312,
        "level": "info",
        "args": ["Execution completed"]
      }
    ],
    "result": {
      "age": 17,
      "name": "Jin Xiaomeng"
    }
  },
  "error": "",
  "status": "finished"
}
```

## Notes

1. Since script tokens allow third parties to access server resources on the platform, for improved **security**, we require you to complete identity verification (can be ignored if already verified)
2. Script Token is the credential for external script execution and belongs to **personal private information**. Through the script token combined with the script webhook, you can read and edit your corresponding files. Please manage it properly and do not share it externally
3. Script Token is bound to the user. Each user can have **at most one** script token. When creating a new token, you need to delete the old token first (the newly created script token must be different from the original token)
4. Script Token expires in 180 days by default. Users can manually extend it when creating, without restrictions, and can extend multiple times
