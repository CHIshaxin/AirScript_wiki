---
title: "MergeCells 属性"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/CellFormat/%E5%B1%9E%E6%80%A7/MergeCells%20%E5%B1%9E%E6%80%A7.html
---
# MergeCells (属性)

如果区域或样式包含合并单元格，则为 **True**。**Variant**。可读写。

## 示例

```javascript
/*此示例显示查找条件中是否包含合并的单元格。*/
function test() {
    console.log(Application.FindFormat.MergeCells)
}
```

```javascript
/*此示例设置查找条件中不包含合并的单元格。*/
function test() {
    Application.FindFormat.MergeCells = false
}
```
