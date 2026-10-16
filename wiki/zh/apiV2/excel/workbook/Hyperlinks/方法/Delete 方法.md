---
title: "Delete (方法)"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/Hyperlinks/%E6%96%B9%E6%B3%95/Delete%20%E6%96%B9%E6%B3%95.html
updated: 1789700916000
hash: Q2xzZiNB
---

# Delete (方法)

删除对象。

## 示例

```javascript
/*本示例取消A1:C10区域的所有超链接*/
function test() {
    Range("A1:C10").Hyperlinks.Delete()
}
```

```javascript
/*本示例取消工作表Sheet1中所有超链接。*/
function test() {
    Application.Worksheets.Item("Sheet1").Hyperlinks.Delete()
}
```
