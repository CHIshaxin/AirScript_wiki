---
title: "AddDataField 方法"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/PivotTable/%E6%96%B9%E6%B3%95/AddDataField%20%E6%96%B9%E6%B3%95.html
---
# AddDataField (方法)

将数据字段添加到数据透视表中。返回一个 **[PivotField](./../../PivotField/PivotField 对象.html)** 对象，该对象表示新的数据字段。

## 参数

| 属性 | 数据类型 | 必填 | 说明 |
| --- | --- | --- | --- |
| Field | IDispatch | 必填 | 服务器上的唯一字段。如果源数据是联机分析处理 (OLAP)（OLAP：为查询和报表（而不是处理事务）而进行了优化的数据库技术。OLAP 数据是按分级结构组织的，它存储在多维数据集而不是表中。），则唯一字段是多维数据集字段。如果源数据不是 OLAP（非 OLAP 源数据（非 OLAP 源数据：非 OLAP 源数据是数据透视表或数据透视图使用的基本数据，该数据来自 OLAP 数据库之外的源。这些数据源包括关系数据库、 ET 工作表中的清单以及文本文件数据库。）），则唯一字段是数据透视表字段。 |
| Caption | any | 可选 | 数据透视表中使用的标签，用于识别该数据字段。 |
| Function | any | 可选 | 在已添加数据字段中执行的函数。 |

## 返回值

PivotField

## 示例

```javascript
/*本示例将标题为“Total Score”的数据字段添加到名为“PivotTable1”的数据透视表中，假定存在包含一个标题为“Score”列的表格。*/
function test() {
    let pTab = ActiveSheet.PivotTables("PivotTable1")
    pTab.AddDataField(ActiveSheet.PivotTables("PivotTable1").PivotFields("Score"), "Total Score")
}
```

```javascript
/*本示例为工作表中第一张数据透视表添加 y 数据字段计数项，假定存在包含一个标题为“y”列的表格。*/
function test() {
    let pTab = ActiveSheet.PivotTables(1)
    pTab.AddDataField(pTab.PivotFields("y"), undefined, xlCount)
}
```
