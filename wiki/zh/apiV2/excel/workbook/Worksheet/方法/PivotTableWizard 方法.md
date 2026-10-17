---
title: "PivotTableWizard (方法)"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/Worksheet/%E6%96%B9%E6%B3%95/PivotTableWizard%20%E6%96%B9%E6%B3%95.html
updated: 1789700916000
hash: CZ3AkyNa
---

# PivotTableWizard (方法)

创建一个新的数据透视表。此方法不显示“数据透视表向导”，它不适用于 OLE DB 数据源。请使用 **Add** 方法添加数据透视表缓存，然后创建基于该缓存的数据透视表。

## 参数

| 属性 | 数据类型 | 必填 | 说明 |
| --- | --- | --- | --- |
| SourceType | any | 可选 | 一个代表报表数据源的 **[XlPivotTableSourceType](./../../枚举/XlPivotTableSourceType 枚举.html)** 值。如果指定了此参数，那么必须同时指定 *SourceData*。如果忽略 *SourceType* 和 *SourceData*，ET 将假定源类型为 **xlDatabase**，并假定源数据来自命名区域“Database”。如果该命名区域不存在，那么当选定区域所在的当前区域中包含数据的单元格超过 10 个时，ET 就使用当前区域。否则，此方法将失败。 |
| SourceData | any | 可选 | 新报表的数据。它可以是一个 **[Range](./../../Range/Range 对象.html)** 对象、一个区域数组或是代表其他报表名称的一个文本常量。对于外部数据库而言，*SourceData* 是一个包含 SQL 查询字符串的字符串数组，其中的每个元素最长为 255 个字符。您应该使用 *Connection* 参数指定 ODBC 连接字符串。为了能与旧版本的 ET 兼容，*SourceData* 可以是一个二元数组。第一个元素是用于指定数据的 ODBC 源的连接字符串，第二个元素是用于获取数据的 SQL 查询字符串。如果指定 *SourceData*，则必须同时指定 *SourceType*。如果活动单元格位于 *SourceData* 区域内，则必须同时指定 *TableDestination*。 |
| TableDestination | any | 可选 | 一个 **[Range](./../../Range/Range 对象.html)** 对象，它指定报表在工作表中的位置。如果省略此参数，则将报表置于活动单元格中。 |
| TableName | any | 可选 | 用于指定新报表名称的字符串。 |
| RowGrand | any | 可选 | 如果为 **True**，则显示报表中的行总计。 |
| ColumnGrand | any | 可选 | 如果为 **True**，则显示报表中的列总计。 |
| SaveData | any | 可选 | 如果为 **True**，则保存报表中的数据。如果为 **False**，则仅保存报表的定义。 |
| HasAutoFormat | any | 可选 | 如果为 **True**，当刷新报表或移动字段时，ET 将自动设置其格式。 |
| AutoPage | any | 可选 | 仅当 *SourceType* 为 **xlConsolidation** 时有效。如果值为 **True**，ET 将为合并创建页字段。如果 *AutoPage* 为 **False**，则必须创建一个或多个页字段。 |
| Reserved | any | 可选 | 不在 ET 中使用。 |
| BackgroundQuery | any | 可选 | 如果为 **True**，则 ET 将异步执行（后台执行）报表查询。默认值为 **False**。 |
| OptimizeCache | any | 可选 | 如果为 **True**，则在构造数据透视表的缓存时对其进行优化。默认值为 **False**。 |
| PageFieldOrder | any | 可选 | 在数据透视表布局中添加页字段的顺序。可为以下 **[XlOrder](./../../枚举/XlOrder 枚举.html)** 常量之一：**xlDownThenOver** 或 **xlOverThenDown**。默认值为 **xlDownThenOver**。 |
| PageFieldWrapCount | any | 可选 | 数据透视表中每列或每行的页字段数。默认值为 0（零）。 |
| ReadData | any | 可选 | 如果为 **True**，则创建数据透视表缓存以包含外部数据库中的所有记录；此时缓存可能会很大。如果 *ReadData* 为 **False**，则可在实际读取数据之前，将某些字段设为基于服务器的页字段。 |
| Connection | any | 可选 | 包含 ODBC 设置的字符串，这些设置使得 ET 可以连接 ODBC 数据源。连接字符串的格式为“ODBC;<连接字符串>”。该参数取代以前为 **PivotCache** 对象的 **Connection** 属性所做的任何设置。 |

## 返回值

一个代表新数据透视表的 **[PivotTable](./../../PivotTable/PivotTable 对象.html)** 对象。

## 示例

```javascript
/*本示例在活动工作表通过 A1:C100 单元格区域创建数据透视表。*/
function test() {
    ActiveSheet.PivotTableWizard(xlDatabase, Range("A1:C100"))
}
```

```javascript
/*本示例通过第二张工作表中 A1:D4 单元格区域在新的工作表中创建数据透视表，并将数据透析表命名为成绩表。*/
function test() {
    let sheet2 = Application.Sheets.Item(2)
    let sheet = ActiveWorkbook.Worksheets.Add()
    sheet2.PivotTableWizard(xlDatabase, sheet2.Range("A1:D4"), sheet.Range("C2"), "成绩表")
}
```
