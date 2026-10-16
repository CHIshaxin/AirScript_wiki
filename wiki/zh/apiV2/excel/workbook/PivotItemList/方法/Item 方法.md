---
title: "Item (方法)"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/PivotItemList/%E6%96%B9%E6%B3%95/Item%20%E6%96%B9%E6%B3%95.html
updated: 1789700916000
hash: Djc8TWiU
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

包含在集合中的一个 PivotItem 对象。

## 示例

```javascript
/*本示例将包含单元格 L4 的数据透视项行轴上第一个数据项设置为不可见。*/
function test() {
    Application.Range("L4").PivotCell.RowItems.Item(1).Visible = false
}
```

```javascript
/*本示例显示包含活动单元格的数据透视项行轴上第二个数据项的值。*/
function test() {
    console.log(ActiveCell.PivotCell.RowItems.Item(2).Value)
}
```
