---
title: "Comment 对象"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/Comment/Comment%20%E5%AF%B9%E8%B1%A1.html
---
# Comment (对象)

代表单元格批注。

## 说明

**[Comment](./../Comment/Comment 对象.html)** 对象是 **[Comments](./../Comments/Comments 对象.html)** 集合的成员。

使用 **Comment** 属性可返回 **[Comment](./../Comment/Comment 对象.html)** 对象。

## 示例

```javascript
/*本示例更改单元格 E5 中的批注文本。*/
function test() {
    Application.Worksheets.Item(1).Range("E5").Comment.Text("reviewed on " + Date())
}
```

使用 **Comments**(*index*)（其中 *index* 为批注号）可返回 **[Comments](./../Comments/Comments 对象.html)** 集合中的单条批注。

```javascript
/*本示例隐藏第一张工作表中的第二条批注。*/
function test() {
    Application.Worksheets.Item(1).Comments.Item(2).Visible = false
}
```

使用 **AddComment** 方法可在区域内添加批注。

```javascript
/*本示例在第一张工作表的单元格 E5 中添加批注。*/
function test() {
    let myComment = Application.Worksheets.Item(1).Range("E5").AddComment()
    myComment.Visible = false
    myComment.Text("reviewed on " + Date())
}
```
