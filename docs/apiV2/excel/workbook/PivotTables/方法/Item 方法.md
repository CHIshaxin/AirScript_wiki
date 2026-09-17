---
title: "Item 方法"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/PivotTables/%E6%96%B9%E6%B3%95/Item%20%E6%96%B9%E6%B3%95.html
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

包含在集合中的一个 **[PivotTable](./../../PivotTable/PivotTable 对象.html)** 对象。

## 示例

```javascript
/*本示例将工作表 Sheet1 中第一个数据透视表上的 age 字段设置为行字段。*/
function test() {
    Worksheets.Item("sheet1").PivotTables().Item(1).PivotFields("age").Orientation = xlRowField
}
```

```javascript
/*本示例显示活动工作表中第一张数据透视表值字段的数量。*/
function test() {
    console.log(ActiveSheet.PivotTables().Item(1).DataFields().Count)
}
```
