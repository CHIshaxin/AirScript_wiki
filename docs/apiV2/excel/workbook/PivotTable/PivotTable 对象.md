---
title: "PivotTable 对象"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/PivotTable/PivotTable%20%E5%AF%B9%E8%B1%A1.html
---
# PivotTable (对象)

代表工作表上的数据透视表。

## 说明

**[PivotTable](./../PivotTable/PivotTable 对象.html)** 对象是 **[PivotTables](./../PivotTables/PivotTables 对象.html)** 集合的成员。 **[PivotTables](./../PivotTables/PivotTables 对象.html)** 集合包含某一张工作表上的所有 **[PivotTable](./../PivotTable/PivotTable 对象.html)** 对象。

因为对数据透视表进行编程可能会很复杂，所以，最方便的做法是将数据透视表操作录制到宏中，然后再修订所录制的宏代码。

使用 **PivotTables** (*index*)（其中 *index* 是数据透视表索引号或名称）可返回一个 **[PivotTable](./../PivotTable/PivotTable 对象.html)** 对象。

## 示例

```javascript
/*本示例使 Sheet1 上第一张数据透视表中的字段“age”成为行字段。*/
function test() {
    Worksheets.Item("Sheet1").PivotTables(1).PivotFields("age").Orientation = xlRowField
}
```

```javascript
/*本示例使第一张工作表上的第一个数据透视表使用合并单元格外部行项、列项、分类汇总和总计的标志。*/
function test() {
    Application.Worksheets.Item(1).PivotTables(1).MergeLabels = true
}
```
