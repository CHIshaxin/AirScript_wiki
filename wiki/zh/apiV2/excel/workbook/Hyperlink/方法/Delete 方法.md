---
title: "Delete (方法)"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/Hyperlink/%E6%96%B9%E6%B3%95/Delete%20%E6%96%B9%E6%B3%95.html
updated: 1789700916000
hash: C6p3HXXk
---

# Delete (方法)

删除对象。

## 示例

```javascript
/*本示例取消单元格A1中的第一项超链接*/
function test() {
    let hl = Range("A1").Hyperlinks.Item(1)
    hl.Delete()
}
```

```javascript
/*本示例取消工作表Sheet1中的第二项超链接*/
function test() {
    let hl = Application.Worksheets.Item("Sheet1").Hyperlinks.Item(2)
    hl.Delete()
}
```
