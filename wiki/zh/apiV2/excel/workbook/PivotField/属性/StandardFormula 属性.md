---
title: "StandardFormula (属性)"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/PivotField/%E5%B1%9E%E6%80%A7/StandardFormula%20%E5%B1%9E%E6%80%A7.html
updated: 1789700916000
hash: C9cOe1Cz
---

# StandardFormula (属性)

返回或设置一个 **String** 值，该值指定使用标准英语（美国）格式的公式。可读写。

## 说明

**StandardFormula** 属性主要影响具有日期或数字格式的项目名称。该属性提供了一种方法可指定或查询给定计算项的公式。

**StandardFormula** 属性是国际通用的，**Formula** 属性则不是。

## 示例

```javascript
/*本示例向“score”字段中添加 10，并将其显示为字段“期中成绩”中的计算项。本示例假定数据透视表位于活动工作表上，并且标题为“score”的字段位于模拟运算表中。*/
function test() {
    let pvtTable = ActiveSheet.PivotTables(1)
    // Change the calculated field by adding "10".
    pvtTable.CalculatedFields().Item("期中成绩").StandardFormula = "score + 10"
}
```

```javascript
/*本示例显示第一张工作表中数据透视表的字段“期中成绩”的公式。*/
function test() {
    let pvtField = Worksheets.Item(1).Range("I1").PivotTable.PivotFields("期中成绩")
    console.log(pvtField.StandardFormula)
}
```
