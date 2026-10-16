---
title: "Item (方法)"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/PivotFields/%E6%96%B9%E6%B3%95/Item%20%E6%96%B9%E6%B3%95.html
updated: 1789700916000
hash: Dm4g658g
---

# Item (方法)

从集合中返回一个对象。

## 说明

对象的文本名称就是 **Name** 和 **Value** 属性的值。

## 参数

| 属性 | 数据类型 | 必填 | 说明 |
| --- | --- | --- | --- |
| Index | any | 必填 | 对象的名称或索引号。 |

## 返回值

一个 Object 值，它代表集合包含的某个对象。

## 示例

```javascript
/*本示例使 Sheet1 上第一张数据透视表中的字段“id”成为行字段。*/
function test() {
    Application.Worksheets.Item("Sheet1").PivotTables(1).PivotFields().Item("id").Orientation = xlRowField
}
```

```javascript
/*本示例在活动工作表的第一个数据透视表的“name”字段之后添加一个分页符。*/
function test() {
    ActiveSheet.PivotTables(1).PivotFields().Item("name").LayoutPageBreak = true
}
```
