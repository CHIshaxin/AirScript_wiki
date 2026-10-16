---
title: "Item (方法)"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/IconCriteria/%E6%96%B9%E6%B3%95/Item%20%E6%96%B9%E6%B3%95.html
updated: 1789700916000
hash: D9TWLZQL
---

# Item (方法)

从 **[IconCriteria](./../../IconCriteria/IconCriteria 对象.html)** 集合中返回单个 **IconCriterion** 对象。只读。

## 说明

*Index* 参数的值不能大于为图标集条件格式设置的条件的数目。条件与为图标集分配的阈值等效。若要查找阈值的数目，请使用 **Count** 属性。

## 参数

| 属性 | 数据类型 | 必填 | 说明 |
| --- | --- | --- | --- |
| Index | any | 必填 | **IconCriterion** 对象的索引号。 |

## 示例

```javascript
/*本示例显示第一张工作表上区域 A1:A10 中第一个图标集条件格式的第一个IconCriterion的类型是否为xlConditionValuePercent。*/
function test() {
    let iconSet = Application.Worksheets.Item(1).Range("A1:A10").FormatConditions.Item(1)
    console.log(iconSet.IconCriteria.Item(1).Type == xlConditionValuePercent)
}
```

```javascript
/*本示例将活动工作表上区域 A1:A10 中第二个图标集条件格式的第二个IconCriterion的图标设置为xlIconYellowExclamation。*/
function test() {
    let iconSet = ActiveSheet.Range("A1:A10").FormatConditions.Item(2)
    iconSet.IconCriteria.Item(2).Icon = xlIconYellowExclamation
}
```
