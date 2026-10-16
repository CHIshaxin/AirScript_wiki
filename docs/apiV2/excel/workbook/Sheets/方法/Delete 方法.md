---
title: "Delete 方法"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/Sheets/%E6%96%B9%E6%B3%95/Delete%20%E6%96%B9%E6%B3%95.html
---
# Delete (方法)

删除对象。

## 示例

```javascript
/*此示例删除活动工作簿中名称为“Sheet1”和“Sheet2”工作表。*/
function test() {
    let names = ["Sheet1", "Sheet2"]
    Application.ActiveWorkbook.Sheets.Item(names).Delete()
}
```

```javascript
/*此示例删除索引为2和3的工作表。*/
function test() {
    Sheets.Item([2, 3]).Delete()
}
```
