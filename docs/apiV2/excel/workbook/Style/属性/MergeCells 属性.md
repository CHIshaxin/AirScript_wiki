---
title: "MergeCells 属性"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/Style/%E5%B1%9E%E6%80%A7/MergeCells%20%E5%B1%9E%E6%80%A7.html
---
# MergeCells (属性)

如果样式包含合并的单元格，则为 **True**。**Variant** 型，可读写。

## 示例

```javascript
/*本示例将A1:B2单元格区域合并为一个单元格。*/
function test() {
    Range("A1:B2").MergeCells = true
}
```

```javascript
/*本示例显示活动工作簿第一个样式是否包含合并的单元格。*/
function test() {
    console.log(ActiveWorkbook.Styles.Item(1).MergeCells)
}
```
