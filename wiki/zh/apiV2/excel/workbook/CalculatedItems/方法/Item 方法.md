---
title: "Item (方法)"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/CalculatedItems/%E6%96%B9%E6%B3%95/Item%20%E6%96%B9%E6%B3%95.html
updated: 1789700916000
hash: BJ4vyf0g
---

# Item (方法)

从集合中返回一个对象。

## 说明

对象的文本名称就是 **Name** 和 **Value** 属性的值。

## 参数

| 属性 | 数据类型 | 必填 | 说明 |
|---|---|---|---|
| Index | any | 必填 | 对象的名称或索引号。 |

## 返回值

包含在集合中的一个 **[PivotItem](./../../PivotItem/PivotItem 对象.html)** 对象。

## 示例

```javascript
/*此示例隐藏第一张工作表上第一张数据透视表中字段“name”的第一个计算项。*/
function test() {
    let pvtItem = Application.Worksheets.Item(1).PivotTables(1).PivotFields("name").CalculatedItems().Item(1)
    pvtItem.Visible = false
}
```

```javascript
/*本示例显示工作表 Sheet1 上第一张数据透视表的字段“name”中名为“公式2”的计算项的公式。*/
function test() {
    let pvtItem = Worksheets.Item("Sheet1").PivotTables(1).PivotFields("name").CalculatedItems().Item("公式2")
    console.log(pvtItem.Formula)
}
```
