---
title: "PivotTables (对象)"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/PivotTables/PivotTables%20%E5%AF%B9%E8%B1%A1.html
updated: 1789700916000
hash: DOdAwhi2
---

# PivotTables (对象)

指定的工作表中所有 **[PivotTable](./../PivotTable/PivotTable 对象.html)** 对象的集合。

## 说明

因为对数据透视表进行编程可能会很复杂，所以，最方便的做法是将数据透视表操作录制到宏中，然后再修订所录制的宏代码。

使用 **PivotTables** 方法可返回 **[PivotTables](./../PivotTables/PivotTables 对象.html)** 集合。

## 示例

```javascript
/*本示例显示工作表 Sheet1 中数据透视表的数量。*/
function test() {
    console.log(Worksheets.Item("Sheet1").PivotTables().Count)
}
```

```javascript
/*本示例基于工作表 Sheet1 中第一张数据透视表的缓存在新工作表上创建新数据透视表。*/
function test() {
    Worksheets.Add().Activate()
    let pvtCache = Worksheets.Item("Sheet1").PivotTables(1).PivotCache()
    ActiveSheet.PivotTables().Add(pvtCache, Range("A1"), "student")
}
```

使用 **PivotTables**(*index*)（其中 *index* 是数据透视表索引号或名称）可返回一个 **[PivotTable](./../PivotTable/PivotTable 对象.html)** 对象。

```javascript
/*本示例将工作表 Sheet1 中第一个数据透视表上的 score 字段设置为行字段。*/
function test() {
    Worksheets.Item("sheet1").PivotTables().Item(1).PivotFields("score").Orientation = xlRowField
}
```
