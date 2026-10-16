---
title: "Item (方法)"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/TableStyles/%E6%96%B9%E6%B3%95/Item%20%E6%96%B9%E6%B3%95.html
updated: 1789700916000
hash: Bo3hsg7l
---

# Item (方法)

从集合中返回一个对象。

## 参数

| 属性 | 数据类型 | 必填 | 说明 |
| --- | --- | --- | --- |

## 返回值

TableStyle

## 示例

```javascript
/*本示例显示活动工作簿的表格样式中第一个样式名称。*/
function test() {
    console.log(ActiveWorkbook.TableStyles.Item(1).Name)
}
```

```javascript
/*本示例显示第一张工作簿的表格样式中第二个样式是否在数据透视表样式的库中显示。*/
function test() {
    let tableStyle = Workbooks.Item(1).TableStyles.Item(2)
    console.log(tableStyle.ShowAsAvailablePivotTableStyle)
}
```
