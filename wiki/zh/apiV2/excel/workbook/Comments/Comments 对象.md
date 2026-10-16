---
title: "Comments (对象)"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/Comments/Comments%20%E5%AF%B9%E8%B1%A1.html
updated: 1789700916000
hash: CK0V6WRA
---

# Comments (对象)

由单元格批注组成的集合。

## 说明

每个批注都由一个 **[Comment](./../Comment/Comment 对象.html)** 对象代表。

使用 **Comments** 属性可返回 **[Comments](./../Comments/Comments 对象.html)** 集合。

## 示例

```javascript
/*本示例隐藏第一张工作表上的所有批注。*/
function test() {
    let cmt = Application.Worksheets.Item(1).Comments
    for (let c = 1; c <= cmt.Count; c++) {
        cmt.Item(c).Visible = false
    }
}
```

使用 **AddComment** 方法可在区域内添加批注。

```javascript
/*本示例在第一张工作表的单元格 E5 上添加一个批注。*/
function test() {
    let myComment = Application.Worksheets.Item(1).Range("E5").AddComment()
    myComment.Visible = false
    myComment.Text("reviewed on " + Date())
}
```

使用 **Comments**(*index*)（其中 *index* 为批注号）可返回 **[Comments](./../Comments/Comments 对象.html)** 集合中的单条批注。

```javascript
/*本示例隐藏第一张工作表上的第二个批注。*/
function test() {
    Application.Worksheets.Item(1).Comments.Item(2).Visible = false
}
```
