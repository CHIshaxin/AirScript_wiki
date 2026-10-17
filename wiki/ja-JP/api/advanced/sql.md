---
title: "データベース API"
source: https://airsheet.wps.cn/docs/ja-JP/api/advanced/sql.html
updated: 1789700916000
hash: rINOEclt
---

# データベース API

AirScript はグローバルな SQL オブジェクトを公開します。開発者はそのプロパティとメソッドを使って**外部データベース**サービスに接続します。接続に成功すると、SQL 文を実行してデータの挿入・更新・削除・照会ができます。

ヒント

SQL でデータベースに接続する前に、スクリプトエディターのツールバー **サービス** メニューから `Database API` サービスを追加してください。

### クイックスタート

```js
// MySQL に接続
const connection = SQL.connect(SQL.Drivers.MySQL, {
  host: '127.0.0.1',
  username: 'root',
  password: '123456',
  database: 'mydb',
  port: 3306
})

// SQL 実行: test から全行を照会
const result1 = connection.queryAll('SELECT * FROM test')
// 結果を出力
console.log(result1)

// SQL 実行: 行を挿入
const result2 = connection.queryAll(
  'INSERT INTO test (id,test_data) VALUES (?,?), (?,?)',
  [1, 1, 2, 2]
)
// 結果を出力
console.log(result2)

// 接続を閉じる
connection.close()
```

### プロパティ

| プロパティ | 型 | 説明 |
| --- | --- | --- |
| [Drivers](#drivers) | object | データベースドライバーの定数 |
| [Types](#types) | object | 列型の定数（SQL Server のみ） |

### メソッド

| メソッド | 戻り値の型 | 説明 |
| --- | --- | --- |
| [connect()](#connect) | Connection | データベースに接続する |
| [Connection.queryAll()](#queryall) | Result | SQL 文を実行する |
| [Connection.close()](#close) | null | 接続を閉じる |

## Drivers

[connect()](#connect) に渡すドライバー定数。読み取り専用。現在 MySQL と SQL Server をサポートします。

### プロパティ

| パラメーター | 型 | 説明 |
| --- | --- | --- |
| MySQL | string | MySQL ドライバー |
| PostgreSQL | string | PostgreSQL ドライバー |
| SQLServer | string | SQL Server ドライバー |

## Types

列型の定数。**備考**: SQL Server にのみ必要。MySQL ではこれらの型は不要です。

### プロパティ

#### 正確な数値型

| プロパティ | 対応する JavaScript の型 |
| --- | --- |
| Bit | Boolean |
| TinyInt | Number |
| SmallInt | Number |
| Int | Number |
| BigInt | String |
| Numeric | Number |
| Decimal | Number |
| SmallMoney | Number |
| Money | Number |

#### 近似数値型

| プロパティ | 対応する JavaScript の型 |
| --- | --- |
| Float | Number |
| Real | Number |

#### 日付と時刻

| プロパティ | 対応する JavaScript の型 |
| --- | --- |
| SmallDateTime | Date |
| DateTime | Date |
| DateTime2 | Date |
| DateTimeOffset | Date |
| Time | Date |
| Date | Date |

#### 文字列

| プロパティ | 対応する JavaScript の型 |
| --- | --- |
| Char | String |
| VarChar | String |
| Text | String |

#### Unicode 文字列

| プロパティ | 対応する JavaScript の型 |
| --- | --- |
| NChar | String |
| NVarChar | String |
| NText | String |

#### バイナリ文字列

| プロパティ | 対応する JavaScript の型 |
| --- | --- |
| Binary | Buffer |
| VarBinary | Buffer |
| Image | Buffer |

#### その他のデータ型

| プロパティ | 対応する JavaScript の型 |
| --- | --- |
| Null | null |
| TVP | Object |
| UDT | Buffer |
| UniqueIdentifier | String |
| Variant | any |
| xml | String |

## connect()

データベースに接続します。現在 MySQL、PostgreSQL、SQL Server をサポートします。SQL を実行するための接続オブジェクトを返します。プログラム終了前に [close()](#close) を呼び出してください。

### パラメーター

| パラメーター | 型 | 既定値 | 必須 | 説明 |
| --- | --- | --- | --- | --- |
| driver | [Driver](#drivers) | null | はい | 対象 DB ドライバー |
| options | [Options](#options) | null | はい | 接続オプション |

### options

| パラメーター | 型 | 既定値 | 必須 | 説明 |
| --- | --- | --- | --- | --- |
| host | string | null | はい | データベースホスト |
| port | number | null | はい | データベースポート |
| username | string | null | はい | ユーザー名 |
| password | string | null | はい | パスワード |
| database | string | null | はい | データベース名 |

### 戻り値

Connection - データベース接続オブジェクト

### 例

```js
// MySQL に接続
const connection = SQL.connect(SQL.Drivers.MySQL, {
  host: '127.0.0.1',
  port: 3340,
  username: 'jinxiaomeng',
  password: '123',
  database: 'WPS_TEST'
})
```

## Connection.queryAll()

[connect()](#connect) が成功したら、返された接続で `queryAll()` を呼び出して SQL を実行します。

### パラメーター

| パラメーター | 型 | 既定値 | 必須 | 説明 |
| --- | --- | --- | --- | --- |
| sql | string | null | はい | 実行する SQL |
| InsertData | any[] \| [InsertData](#insertdata) | null | いいえ | INSERT のバインド値 |

### InsertData

| パラメーター | 型 | 既定値 | 必須 | 説明 |
| --- | --- | --- | --- | --- |
| name | string | null | はい | 列名 |
| value | string | null | はい | 値 |
| type | [Types](#types) | null | いいえ | 列型。SQL Server の INSERT では必須 |

### 戻り値

影響行数と返却行を含む Result オブジェクト。

| プロパティ | 型 | 説明 |
| --- | --- | --- |
| affectRowCount | number | SQL 文によって影響を受けた行数 |
| rows | Array | 照会した表構造に一致する結果行 |

### 戻り値の形の例

```json
// SELECT の結果
{
  "affectRowCount": 0,
  "rows": [
    [
      {
        "name": "1",
        "value": 2
      }
    ]
  ]
}

// INSERT/UPDATE/DELETE の結果
{
  "affectRowCount": 1,
  "rows": []
}
```

### 例

```js
// SQL Server に接続
const connection = SQL.connect(SQL.Drivers.SQLServer, {
  host: 'x.x.x.x',
  username: 'x',
  password: 'x',
  database: 'x',
  port: 1433
})

// 2 行 INSERT
const result1 = connection.queryAll(
  'INSERT INTO TestSchema.Employees (Name, Location) OUTPUT INSERTED.Id VALUES (@Name, @Location);',
  [
    {
      name: 'Name',
      type: SQL.Types.NVarChar,
      value: 'zhangsan'
    },
    {
      name: 'Location',
      type: SQL.Types.NVarChar,
      value: 'zhuhai'
    }
  ]
)

// 結果を出力
console.log(result1)

// 従業員を SELECT
const result2 = connection.queryAll(
  'SELECT Id, Name, Location FROM TestSchema.Employees;'
)

// 結果を出力
console.log(result2)

// 接続を閉じる
connection.close()
```

## Connection.close()

データベース接続を閉じます。プログラム終了前に呼び出してください。

### 例

```js
// MySQL に接続
const connection = SQL.connect(SQL.Drivers.MySQL, {
  host: '127.0.0.1',
  port: 3340,
  username: 'jinxiaomeng',
  password: '123',
  database: 'WPS_TEST'
})

// 何か処理する

// 接続を閉じる
connection.close()
```
