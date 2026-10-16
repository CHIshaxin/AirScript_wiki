---
title: "Formula (属性)"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/PivotField/%E5%B1%9E%E6%80%A7/Formula%20%E5%B1%9E%E6%80%A7.html
updated: 1789700916000
hash: B8EZ11id
---

# Formula (属性)

返回或设置一个 **String** 值，它代表 A1 样式表示法和宏语言中的对象的公式。

## 说明

此属性对于 OLAP（OLAP：为查询和报表（而不是处理事务）而进行了优化的数据库技术。OLAP 数据是按分级结构组织的，它存储在多维数据集而不是表中。） 数据源无效。

如果单元格包含一个常量，此属性返回该常量。如果单元格为空，此属性返回空字符串。如果单元格包含公式，**Formula** 属性将该公式作为字符串返回，所用格式与在编辑栏（包括等号）中显示时的格式相同。

如果将单元格的值或者公式设置为日期类型，则 ET 将检查此单元格的数字格式是否符合日期或者时间格式。如果不符合，ET 将把数字格式设置为默认的短日期格式。

如果指定区域是一维或二维区域，则可将公式指定为 Javascript 中相同维数的数组。同样，也可在 Javascript 数组中使用公式。

如果为多单元格区域设置公式，则会用公式填充该区域所有的单元格。

## 示例

```javascript
/*本示例显示活动工作表 Sheet1 中数据透视表的字段“期中成绩”的公式。*/
function test() {
    let pvtField = ActiveSheet.Range("I3").PivotTable.PivotFields("期中成绩")
    console.log(pvtField.Formula)
}
```

```javascript
/*本示例设置工作表 Sheet1 中第一个数据透视表的字段“期末成绩”的公式。*/
function test() {
    let pvtField = Worksheets.Item("Sheet1").PivotTables(1).PivotFields("期末成绩")
    pvtField.Formula = "=grades*0.6"
}
```
