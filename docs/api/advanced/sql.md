---
title: "数据库 API"
source: https://airsheet.wps.cn/docs/api/advanced/sql.html
---
# 数据库 API

AirScript 提供一个全局的 SQL 对象，开发者可通过此对象提供的属性和方法连接到**外部数据库**服务，连接成功后即可执行 SQL 语句，对数据进行增删改查。

> **提示**
>
> 在使用 SQL 对象连接数据库之前，确保您已添加`数据库API`服务，在脚本编辑器的【工具栏】-【服务】菜单内添加即可。

### 快速使用

```js
// 连接MySQL数据库
const connection = SQL.connect(SQL.Drivers.MySQL, {
  host: '127.0.0.1',
  username: 'root',
  password: '123456',
  database: 'mydb',
  port: 3306
})

// 执行SQL语句，查询test表的所有数据
const result1 = connection.queryAll('SELECT * FROM test')
// 打印执行结果
console.log(result1)

// 执行SQL语句，插入数据
const result2 = connection.queryAll(
  'INSERT INTO test (id,test_data) VALUES (?,?), (?,?)',
  [1, 1, 2, 2]
)
// 打印执行结果
console.log(result2)

// 关闭数据库连接
connection.close()
```

### 属性列表

| 属性名 | 数据类型 | 说明 |
| --- | --- | --- |
| Drivers | object | 数据库连接驱动集 |
| Types | object | 数据库字段类型集（仅适用于 SQL server） |

### 方法列表

| 方法 | 返回类型 | 说明 |
| --- | --- | --- |
| connect() | Connection | 连接目标数据库 |
| Connection.queryAll() | Result | 执行 SQL 语句 |
| Connection.close() | null | 关闭数据库连接 |

## Drivers

数据库驱动集，调用connect()方法连接数据库时传入对应驱动，目前仅支持 MySQL 和 SQL server 两种驱动，只读

### 属性说明

| 属性名 | 数据类型 | 说明 |
| --- | --- | --- |
| MySQL | string | MySQL 数据库驱动 |
| PostgreSQL | string | PostgreSQL 数据库驱动 |
| SQLServer | string | SQL server 数据库驱动 |

## Types

数据库字段类型集，请注意，该类型集仅适用于 SQL server 数据库，MySQL 数据库不需要传递此值

### 属性说明

#### Exact numerics

| 属性名 | 对应 Javascript 类型 |
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

#### Approximate numerics

| 属性名 | 对应 Javascript 类型 |
| --- | --- |
| Float | Number |
| Real | Number |

#### Date and Time

| 属性名 | 对应 Javascript 类型 |
| --- | --- |
| SmallDateTime | Date |
| DateTime | Date |
| DateTime2 | Date |
| DateTimeOffset | Date |
| Time | Date |
| Date | Date |

#### Character Strings

| 属性名 | 对应 Javascript 类型 |
| --- | --- |
| Char | String |
| VarChar | String |
| Text | String |

#### Unicode Strings

| 属性名 | 对应 Javascript 类型 |
| --- | --- |
| NChar | String |
| NVarChar | String |
| NText | String |

#### Binary Strings

| 属性名 | 对应 Javascript 类型 |
| --- | --- |
| Binary | Buffer |
| VarBinary | Buffer |
| Image | Buffer |

#### Other Data Types

| 属性名 | 对应 Javascript 类型 |
| --- | --- |
| Null | null |
| TVP | Object |
| UDT | Buffer |
| UniqueIdentifier | String |
| Variant | any |
| xml | String |

## connect()

连接目标数据库，目前仅支持 MySQL、PostgreSQL 和 SQL server 三种类型的数据库，连接成功后会返回数据库连接对象，可通过此对象执行 SQL 语句，程序结束之前请调用close()方法关闭数据库连接。

### 参数

| 属性 | 数据类型 | 默认值 | 必填 | 说明 |
| --- | --- | --- | --- | --- |
| driver | Driver | null | 是 | 指定目标数据库驱动 |
| options | Options | null | 是 | 数据库连接信息 |

### options 对象说明

| 属性 | 数据类型 | 默认值 | 必填 | 说明 |
| --- | --- | --- | --- | --- |
| host | string | null | 是 | 目标数据库主机 |
| port | number | null | 是 | 目标数据库端口 |
| username | string | null | 是 | 目标数据库连接用户名 |
| password | string | null | 是 | 目标数据库连接密码 |
| database | string | null | 是 | 目标数据库名 |

### 返回值

Connection - 数据库连接对象

### 示例

```js
// 连接MySQL数据库
const connection = SQL.connect(SQL.Drivers.MySQL, {
  host: '127.0.0.1',
  port: 3340,
  username: 'jinxiaomeng',
  password: '123',
  database: 'WPS_TEST'
})
```

## Connection.queryAll()

通过上述的connect()方法成功连接数据库后，会返回数据库连接对象，通过此对象即可调用 queryAll()方法执行 SQL 语句

### 参数

| 属性 | 数据类型 | 默认值 | 必填 | 说明 |
| --- | --- | --- | --- | --- |
| sql | string | null | 是 | 要执行的 sql 语句 |
| InsertData | any[] \| InsertData | null | 否 | 需要插入的数据 |

### InsertData 对象说明

| 属性 | 数据类型 | 默认值 | 必填 | 说明 |
| --- | --- | --- | --- | --- |
| name | string | null | 是 | 插入数据的字段名 |
| value | string | null | 是 | 插入数据的值 |
| type | Types | null | 否 | 插入数据的类型，SQL server 数据库必须传递该类型 |

### 返回值

Result 对象，包含受影响的行数以及返回的数据行

| 属性 | 数据类型 | 说明 |
| --- | --- | --- |
| affectRowCount | number | 执行 sql 语句后受到影响的行数 |
| rows | Array | 数据行，根据实际查询的表的数据结构返回 |

### 返回示例

```json
// 查询时的返回
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

// 增删改时的返回
{
  "affectRowCount": 1,
  "rows": []
}
```

### 示例

```js
// 连接SQL server数据库
const connection = SQL.connect(SQL.Drivers.SQLServer, {
  host: 'x.x.x.x',
  username: 'x',
  password: 'x',
  database: 'x',
  port: 1433
})

// 执行sql语句，插入两条数据
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

// 打印执行结果
console.log(result1)

// 执行sql语句，查询员工表
const result2 = connection.queryAll(
  'SELECT Id, Name, Location FROM TestSchema.Employees;'
)

// 打印执行结果
console.log(result2)

// 关闭数据库连接
connection.close()
```

## Connection.close()

关闭数据库连接，请务必在程序结束前调用此方法

### 示例

```js
// 连接MySQL数据库
const connection = SQL.connect(SQL.Drivers.MySQL, {
  host: '127.0.0.1',
  port: 3340,
  username: 'jinxiaomeng',
  password: '123',
  database: 'WPS_TEST'
})

// do something

// 关闭数据库连接
connection.close()
```
