---
title: "Delete 方法"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/AllowEditRange/%E6%96%B9%E6%B3%95/Delete%20%E6%96%B9%E6%B3%95.html
---
# Delete (方法)

删除对象。

## 示例

```javascript
/*本示例将第一张工作表上第一个 AllowEditRange 对象删除。*/
function test() {
    Worksheets.Item(1).Protection.AllowEditRanges.Item(1).Delete()
}
```

```javascript
/*本示例将活动工作表上第二个 AllowEditRange 对象删除。*/
function test() {
    ActiveSheet.Protection.AllowEditRanges.Item(2).Delete()
}
```
