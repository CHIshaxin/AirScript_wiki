---
title: "NumberFormat (属性)"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/PivotField/%E5%B1%9E%E6%80%A7/NumberFormat%20%E5%B1%9E%E6%80%A7.html
updated: 1789700916000
hash: C_P7_qng
---

# NumberFormat (属性)

返回或设置一个 **String** 值，它代表对象的格式代码。

## 说明

只能对数据字段设置 **NumberFormat** 属性。

格式代码与 **“设置单元格格式”** 对话框中的 **“格式代码”** 选项是同一个字符串。**Format** 函数使用的格式代码字符串与 **NumberFormat** 和 **NumberFormatLocal** 属性使用的格式代码字符串不同。

## 示例

```javascript
/*本示例显示第一张工作表中数据透视表的字段“id”的格式代码。*/
function test() {
    let pvtField = Worksheets.Item(1).Range("I1").PivotTable.PivotFields("id")
    console.log(pvtField.NumberFormat)
}
```

```javascript
/*本示例设置活动工作表中第一张数据透视表的字段“score”的格式代码。*/
function test() {
    ActiveSheet.PivotTables(1).PivotFields("score").NumberFormat = "0.00"
}
```
