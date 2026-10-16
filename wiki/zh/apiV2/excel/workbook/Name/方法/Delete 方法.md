---
title: "Delete (方法)"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/Name/%E6%96%B9%E6%B3%95/Delete%20%E6%96%B9%E6%B3%95.html
updated: 1789700916000
hash: edpM69K8
---

# Delete (方法)

删除对象。

## 示例

```javascript
/*本示例从活动工作簿中删除名称“test”。*/
function test() {
    ActiveWorkbook.Names.Item("test").Delete()
}
```

```javascript
/*本示例从 Sheet1 中删除名称“test1”。*/
function test() {
    Application.Worksheets.Item("Sheet1").Names.Item("test1").Delete()
}
```
