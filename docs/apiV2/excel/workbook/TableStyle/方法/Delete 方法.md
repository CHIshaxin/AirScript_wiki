---
title: "Delete 方法"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/TableStyle/%E6%96%B9%E6%B3%95/Delete%20%E6%96%B9%E6%B3%95.html
---
# Delete (方法)

删除 **[TableStyle](./../../TableStyle/TableStyle 对象.html)** 对象。

## 返回值

void

## 示例

```javascript
/*本示例删除活动工作簿上的表格样式中的最后一个样式。*/
function test() {
    let tableStyles = ActiveWorkbook.TableStyles
    tableStyles.Item(tableStyles.Count).Delete()
}
```

```javascript
/*本示例删除第一张工作簿中指定索引的表格样式。*/
function test() {
    Workbooks.Item(1).TableStyles.Item(285).Delete()
}
```
