---
title: "Delete 方法"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/FormatConditions/%E6%96%B9%E6%B3%95/Delete%20%E6%96%B9%E6%B3%95.html
---
# Delete (方法)

删除对象。

## 示例

```javascript
/*本示例删除活动工作表上单元格区域 A1:A10 中所有条件格式。*/
function test() {
    ActiveSheet.Range("A1:A10").FormatConditions.Delete()
}
```

```javascript
/*本示例删除活动工作表上单元格区域 A1:A10 中所有条件格式，并显示删除前后的条件格式的数量。*/
function test() {
    let formatConditions = ActiveSheet.Range("A1:A10").FormatConditions
    console.log(`删除前条件格式的数量：${formatConditions.Count}`)
    formatConditions.Delete()
    console.log(`删除后条件格式的数量：${formatConditions.Count}`)
}
```
