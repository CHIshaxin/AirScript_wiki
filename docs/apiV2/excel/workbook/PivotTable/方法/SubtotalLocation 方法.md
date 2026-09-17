---
title: "SubtotalLocation 方法"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/PivotTable/%E6%96%B9%E6%B3%95/SubtotalLocation%20%E6%96%B9%E6%B3%95.html
---
# SubtotalLocation (方法)

此方法可更改所有现有透视字段的分类汇总位置。更改分类汇总位置仅对大纲形式的字段产生直接的视觉效果，但是也会对表格形式的字段设置该方法。

## 说明

**SubtotalLocation** 方法会自动设置所有现有透视字段的 **LayoutSubtotalLocation** 属性。

## 参数

| 属性 | 数据类型 | 必填 | 说明 |
| --- | --- | --- | --- |
| Location | XlSubtototalLocationType | 必填 | xlSubtotalLocationType 既可以是 **xlAtTop**，也可以是 **xlAtBottom**。 |

## 示例

```javascript
/*本示例将数据透视表的分类汇总显示在顶部*/
function test() {
    Worksheets.Item(1).PivotTables(1).SubtotalLocation(xlAtTop)
}
```

```javascript
/*本示例将活动工作表中所有数据透视表的分类汇总显示在底部*/
function test() {
    for (i = 1; i <= ActiveSheet.PivotTables().Count; i++) {
        ActiveSheet.PivotTables(i).SubtotalLocation(xlAtBottom)
        console.log(`第${i}张数据透视表的分类汇总已显示在底部`)
    }
}
```
