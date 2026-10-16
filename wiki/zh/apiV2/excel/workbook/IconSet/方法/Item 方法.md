---
title: "Item (方法)"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/IconSet/%E6%96%B9%E6%B3%95/Item%20%E6%96%B9%E6%B3%95.html
updated: 1789700916000
hash: Qxt5LHBj
---

# Item (方法)

返回一个代表图标集中的单个图标的 **[Icon](./../../Icon/Icon 对象.html)** 对象。只读。

## 说明

*Index* 参数的值不能大于图标集中图标的数目。要查找图标集中图标的总数，请使用 **IconSet.Count** 属性。

## 参数

| 属性 | 数据类型 | 必填 | 说明 |
| --- | --- | --- | --- |
| Index | any | 必填 | **[Icon](./../../Icon/Icon 对象.html)** 对象的索引号。 |

## 示例

```javascript
/*本示例显示第一张工作表上区域 B1:B10 中第一个图标集条件格式的图标集中第一个图标的索引。*/
function test() {
    let iconSet = Application.Worksheets.Item(1).Range("B1:B10").FormatConditions.Item(1).IconSet
    console.log(iconSet.Item(1).Index)
}
```

```javascript
/*本示例显示活动工作表上区域 B1:B10 中第二个图标集条件格式的图标集中第一个图标的实际类型。*/
function test() {
    let iconSet = ActiveSheet.Range("B1:B10").FormatConditions.Item(2).IconSet
    console.log(iconSet.Item(1).constructor.name)
}
```
