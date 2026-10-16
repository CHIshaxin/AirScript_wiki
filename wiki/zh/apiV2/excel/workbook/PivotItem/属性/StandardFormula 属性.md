---
title: "StandardFormula (属性)"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/PivotItem/%E5%B1%9E%E6%80%A7/StandardFormula%20%E5%B1%9E%E6%80%A7.html
updated: 1789700916000
hash: CJvu9daS
---

# StandardFormula (属性)

返回或设置一个 **String** 值，该值指定使用标准英语（美国）格式的公式。可读写。

## 说明

**StandardFormula** 属性主要影响具有日期或数字格式的项目名称。该属性提供了一种方法可指定或查询给定计算项的公式。

**StandardFormula** 属性是国际通用的，**Formula** 属性则不是。

## 示例

```javascript
/*本示例向“name”字段中第十个数据项添加 10。本示例假定数据透视表位于活动工作簿上，并且标题为“name”的字段位于模拟运算表中。*/
function test() {
    let pvtTable = ActiveSheet.PivotTables(1)
    pvtTable.PivotFields().Item("name").PivotItems(10).StandardFormula = "=张三+10"
}
```

```javascript
/*本示例显示工作表 Sheet1 上第一张数据透视表的字段“name”中名为“公式1”的数据项的公式。*/
function test() {
    let pvtItem = Worksheets.Item("Sheet1").PivotTables(1).PivotFields("name").PivotItems("公式1")
    console.log(pvtItem.StandardFormula)
}
```
