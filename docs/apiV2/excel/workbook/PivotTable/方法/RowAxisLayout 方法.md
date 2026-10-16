---
title: "RowAxisLayout 方法"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/PivotTable/%E6%96%B9%E6%B3%95/RowAxisLayout%20%E6%96%B9%E6%B3%95.html
---
# RowAxisLayout (方法)

此方法用于同时为所有现有的透视字段设置版式选项。

## 说明

此方法是基本的，因此，它确保了如果不能在任何透视字段上设置版式选项，则这些字段中没有一个的版式选项将会改变，而且不会对数据透视表进行任何更改。

## 参数

| 属性 | 数据类型 | 必填 | 说明 |
| --- | --- | --- | --- |
| RowLayout | XlLayoutRowType | 必填 | 可以是 **xlCompactRow**、**xlTabularRow** 或 **xlOutlineRow**。 |

## 示例

```javascript
/*本示例将第一张数据透视表的版式行类型设置为表格行*/
function test() {
    Worksheets.Item(1).PivotTables(1).RowAxisLayout(xlTabularRow)
}
```

```javascript
/*本示例将第一张数据透视表的版式行类型设置为压缩行，并弹窗提示*/
function test() {
    ActiveSheet.PivotTables(1).RowAxisLayout(xlCompactRow)
    console.log("数据透视表的版式行类型为压缩行")
}
```
