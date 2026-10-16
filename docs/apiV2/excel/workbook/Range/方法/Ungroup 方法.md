---
title: "Ungroup 方法"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/Range/%E6%96%B9%E6%B3%95/Ungroup%20%E6%96%B9%E6%B3%95.html
---
# Ungroup (方法)

在分级显示中对一个区域进行升级（即降低其分级显示的级别）。指定区域必须是行或列，或者行区域或列区域。如果指定区域在数据透视表中，本方法将对该区域内的项取消分组。

## 说明

如果当前单元格在父字段的字段标题中，则该字段中的所有分组将被取消，并且该字段从数据透视表中删除。当父字段中的最后一个分组被取消后，整个字段将从数据透视表中删除。

## 返回值

Variant

## 示例

```javascript
/*此示例取消 ORDER_DATE 字段的分组。*/
function test() {
    let pvtTable = Worksheets.Item("Sheet1").Range("A3").PivotTable
    let groupRange = pvtTable.PivotFields("ORDER_DATE").DataRange
    groupRange.Cells(1).Ungroup()
}
```

```javascript
/*本示例降低 I1:K1 单元格区域分组显示的级别*/
function test() {
    Range("I1:K1").Ungroup()
}
```
