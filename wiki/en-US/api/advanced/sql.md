---
title: "Database API"
source: https://airsheet.wps.cn/docs/en-US/api/advanced/sql.html
updated: 1789700916000
hash: BeD316Kc
---

# Database API

AirScript exposes a global SQL object. Developers use its properties and methods to connect to an **external database** service. After a successful connection, you can run SQL statements to insert, update, delete, and query data.

Tip

Before using SQL to connect to a database, add the `Database API` service from the script editor toolbar **Services** menu.

### Quick start

```js
// Connect to MySQL
const connection = SQL.connect(SQL.Drivers.MySQL, {
  host: '127.0.0.1',
  username: 'root',
  password: '123456',
  database: 'mydb',
  port: 3306
})

// Run SQL: query all rows from test
const result1 = connection.queryAll('SELECT * FROM test')
// Print result
console.log(result1)

// Run SQL: insert rows
const result2 = connection.queryAll(
  'INSERT INTO test (id,test_data) VALUES (?,?), (?,?)',
  [1, 1, 2, 2]
)
// Print result
console.log(result2)

// Close connection
connection.close()
```

### Properties

| Property | Type | Description |
|---|---|---|
| [Drivers](#drivers) | object | Database driver constants |
| [Types](#types) | object | Column type constants (SQL Server only) |

### Methods

| Method | Return type | Description |
|---|---|---|
| [connect()](#connect) | Connection | Connects to a database |
| [Connection.queryAll()](#queryall) | Result | Runs SQL statements |
| [Connection.close()](#close) | null | Closes the connection |

## Drivers

Driver constants passed to [connect()](#connect). Read-only. Currently supports MySQL and SQL Server.

### Properties

| Name | Type | Description |
|---|---|---|
| MySQL | string | MySQL driver |
| PostgreSQL | string | PostgreSQL driver |
| SQLServer | string | SQL Server driver |

## Types

Column type constants. Note: only needed for SQL Server; MySQL does not require these types.

### Properties

#### Exact numerics

| Property | Maps to JavaScript type |
|---|---|
| Bit | Boolean |
| TinyInt | Number |
| SmallInt | Number |
| Int | Number |
| BigInt | String |
| Numeric | Number |
| Decimal | Number |
| SmallMoney | Number |
| Money | Number |

#### Approximate numerics

| Property | Maps to JavaScript type |
|---|---|
| Float | Number |
| Real | Number |

#### Date and Time

| Property | Maps to JavaScript type |
|---|---|
| SmallDateTime | Date |
| DateTime | Date |
| DateTime2 | Date |
| DateTimeOffset | Date |
| Time | Date |
| Date | Date |

#### Character Strings

| Property | Maps to JavaScript type |
|---|---|
| Char | String |
| VarChar | String |
| Text | String |

#### Unicode Strings

| Property | Maps to JavaScript type |
|---|---|
| NChar | String |
| NVarChar | String |
| NText | String |

#### Binary Strings

| Property | Maps to JavaScript type |
|---|---|
| Binary | Buffer |
| VarBinary | Buffer |
| Image | Buffer |

#### Other Data Types

| Property | Maps to JavaScript type |
|---|---|
| Null | null |
| TVP | Object |
| UDT | Buffer |
| UniqueIdentifier | String |
| Variant | any |
| xml | String |

## connect()

Connects to a database. Currently supports MySQL, PostgreSQL, and SQL Server. Returns a connection object used to run SQL. Call [close()](#close) before the program ends.

### Parameters

| Name | Type | Default | Required | Description |
|---|---|---|---|---|
| driver | [Driver](#drivers) | null | Yes | Target DB driver |
| options | [Options](#options) | null | Yes | Connection options |

### options

| Name | Type | Default | Required | Description |
|---|---|---|---|---|
| host | string | null | Yes | Database host |
| port | number | null | Yes | Database port |
| username | string | null | Yes | Username |
| password | string | null | Yes | Password |
| database | string | null | Yes | Database name |

### Return value

Connection - Database connection object

### Example

```js
// Connect to MySQL
const connection = SQL.connect(SQL.Drivers.MySQL, {
  host: '127.0.0.1',
  port: 3340,
  username: 'jinxiaomeng',
  password: '123',
  database: 'WPS_TEST'
})
```

## Connection.queryAll()

After [connect()](#connect) succeeds, use the returned connection to call `queryAll()` and run SQL.

### Parameters

| Name | Type | Default | Required | Description |
|---|---|---|---|---|
| sql | string | null | Yes | SQL to run |
| InsertData | any[] \| [InsertData](#insertdata) | null | No | Insert bind values |

### InsertData

| Name | Type | Default | Required | Description |
|---|---|---|---|---|
| name | string | null | Yes | Column name |
| value | string | null | Yes | Value |
| type | [Types](#types) | null | No | Column type; required for SQL Server inserts |

### Return value

Result object with affected row count and returned rows.

| Property | Type | Description |
|---|---|---|
| affectRowCount | number | Rows affected by the SQL statement |
| rows | Array | Result rows matching the queried table structure |

### Return shape example

```json
// SELECT result
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

// INSERT/UPDATE/DELETE result
{
  "affectRowCount": 1,
  "rows": []
}
```

### Example

```js
// Connect to SQL Server
const connection = SQL.connect(SQL.Drivers.SQLServer, {
  host: 'x.x.x.x',
  username: 'x',
  password: 'x',
  database: 'x',
  port: 1433
})

// INSERT two rows
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

// Print result
console.log(result1)

// SELECT employees
const result2 = connection.queryAll(
  'SELECT Id, Name, Location FROM TestSchema.Employees;'
)

// Print result
console.log(result2)

// Close connection
connection.close()
```

## Connection.close()

Closes the database connection. Call this before the program ends.

### Example

```js
// Connect to MySQL
const connection = SQL.connect(SQL.Drivers.MySQL, {
  host: '127.0.0.1',
  port: 3340,
  username: 'jinxiaomeng',
  password: '123',
  database: 'WPS_TEST'
})

// do something

// Close connection
connection.close()
```
