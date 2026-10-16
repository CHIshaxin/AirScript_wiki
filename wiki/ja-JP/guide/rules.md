---
title: "スクリプト言語"
source: https://airsheet.wps.cn/docs/ja-JP/guide/rules.html
updated: 1789700916000
hash: KuVs2oQL
---

# スクリプト言語

AirScript は標準の [JavaScript](https://developer.mozilla.org/ja/docs/Learn/Getting_started_with_the_web/JavaScript_basics) をスクリプトに使用し、ES6 のほとんどの構文をサポートしています。

## 組み込みグローバルオブジェクト

#### ドキュメント OpenAPI オブジェクト `Application`

Application はドキュメントの OpenAPI オブジェクトです。WPS Docs のドキュメント種別に応じて異なる Application オブジェクトのプロパティが生成され、コードエディター内で直接参照できます。

#### スクリプトコンテキストオブジェクト `Context`

Context はスクリプトのランタイムコンテキストオブジェクトです。`view` や `configuration` などによって生成される、スクリプト実行に関連するプロパティの一部をマウントする役割を持ち、コードエディター内で直接参照してアプリケーション全体の構築を容易にします。

## 非対応の構文

### クラス

```javascript
class C{
    private a; // 予期しないトークン
    public b(){
        return this.a;
    }
}
```

### オブジェクト内でのメソッドの直接定義

```javascript
let obj = {
    sayHello(){ // 予期しないトークン
        console.log("hello")
    }
}
obj.sayHello()
```

### import, export

```javascript
import os from 'os' // 構文エラー
export default os // 構文エラー
```

### ?.

```javascript
function f(argv){
    console.log(argv?.toString()) // 予期しないトークン
}
```

### await

```javascript
async function f() {
    await new Promise((resolve)=>{ // 予期しないトークン
        resolve("")
    })
}
```

### yield

```javascript
function* foo(index) {
  while (index < 2) {
    yield index; // 予期しないトークン
    index++;
  }
}

const iterator = foo(0);

console.log(iterator.next().value);

console.log(iterator.next().value);
```
