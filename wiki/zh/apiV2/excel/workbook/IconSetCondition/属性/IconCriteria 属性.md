---
title: "IconCriteria (属性)"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/IconSetCondition/%E5%B1%9E%E6%80%A7/IconCriteria%20%E5%B1%9E%E6%80%A7.html
updated: 1789700916000
hash: CCT3hm1A
---

# IconCriteria (属性)

返回一个 **[IconCriteria](./../../IconCriteria/IconCriteria 对象.html)** 集合，该集合代表图标集条件格式规则的一组条件。

## 说明

在您具有 **[IconCriteria](./../../IconCriteria/IconCriteria 对象.html)** 集合之后，可以使用该集合的 **Item** 属性返回 **IconCriterion** 对象。**IconCriterion** 对象代表图标集的单个条件。

## 示例

```javascript
/*本示例显示第一张工作表上区域 A1:A10 中第一个图标集条件格式的第一个IconCriterion的类型是否为xlConditionValuePercent。*/
function test() {
    let iconSet = Application.Worksheets.Item(1).Range("A1:A10").FormatConditions.Item(1)
    console.log(iconSet.IconCriteria.Item(1).Type == xlConditionValuePercent)
}
```

```javascript
/*本示例设置活动工作表上区域 A1:A10 中第二个图标集条件格式的第二个IconCriterion的图标为xlIconYellowExclamation。*/
function test() {
    let iconSet = ActiveSheet.Range("A1:A10").FormatConditions.Item(2)
    iconSet.IconCriteria.Item(2).Icon = xlIconYellowExclamation
}
```
