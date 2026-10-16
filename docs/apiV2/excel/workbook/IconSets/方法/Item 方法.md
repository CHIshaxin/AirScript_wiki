---
title: "Item 方法"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/IconSets/%E6%96%B9%E6%B3%95/Item%20%E6%96%B9%E6%B3%95.html
---
# Item (方法)

从 **[IconSets](./../../IconSets/IconSets 对象.html)** 集合中返回单个 **[IconSet](./../../IconSet/IconSet 对象.html)** 对象。只读。

## 说明

*Index* 参数的值不能大于可用的图标集数目。要查找可用于工作簿的图标集的数目，请使用 **Count** 属性。

## 参数

| 属性 | 数据类型 | 必填 | 说明 |
| --- | --- | --- | --- |
| Index | any | 必填 | **[IconSet](./../../IconSet/IconSet 对象.html)** 对象的索引号。 |

## 示例

```javascript
/*本示例显示活动工作簿上第一个图标集中图标的数目。*/
function test() {
    console.log(ActiveWorkbook.IconSets.Item(1).Count)
}
```

```javascript
/*本示例显示第一张工作簿上第二个图标集的名称是否为三色旗。*/
function test() {
    let iconSet = Workbooks.Item(1).IconSets.Item(2)
    console.log(iconSet.ID == xl3Flags)
}
```
