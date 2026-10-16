---
title: "Item (方法)"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/CalculatedFields/%E6%96%B9%E6%B3%95/Item%20%E6%96%B9%E6%B3%95.html
updated: 1789700916000
hash: BRD1yJz_
---

# Item (方法)

从集合中返回一个对象。

## 说明

对象的文本名称是 **Name** 和 **Value** 属性的值。

## 参数

| 属性 | 数据类型 | 必填 | 说明 |
|---|---|---|---|
| Index | any | 必填 | 对象的名称或索引号。 |

## 返回值

集合包含的 **[PivotField](./../../PivotField/PivotField 对象.html)** 对象。

## 示例

```javascript
/*此示例为第一张工作表中第一张数据透视表的第一个计算字段设置公式。*/
function test() {
    Application.Worksheets.Item(1).PivotTables(1).CalculatedFields().Item(1).Formula = "=score*0.6"
}
```

```javascript
/*此示例删除活动工作表中数据透视表的第一个计算字段。*/
function test() {
    ActiveSheet.Range("I1").PivotTable.CalculatedFields().Item(1).Delete()
}
```
