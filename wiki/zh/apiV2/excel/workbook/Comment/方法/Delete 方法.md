---
title: "Delete (方法)"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/Comment/%E6%96%B9%E6%B3%95/Delete%20%E6%96%B9%E6%B3%95.html
updated: 1789700916000
hash: BA0As_Ev
---

# Delete (方法)

删除对象。

## 示例

```javascript
/*本示例删除 A1 单元格的批注。*/
function test() {
    Range("A1").Comment.Delete()
}
```

```javascript
/*本示例遍历活动工作表上所有批注，并删除该工作表中由 jack 添加的所有批注。*/
function test() {
    let comments = Application.ActiveSheet.Comments
    for (let c = 1; c <= comments.Count; c++) {
        if (comments.Item(c).Author == "jack") {
            comments.Item(c).Delete()
        }
    }
}
```
