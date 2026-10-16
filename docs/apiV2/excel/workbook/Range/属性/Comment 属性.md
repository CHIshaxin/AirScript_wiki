---
title: "Comment 属性"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/Range/%E5%B1%9E%E6%80%A7/Comment%20%E5%B1%9E%E6%80%A7.html
---
# Comment (属性)

返回一个 **[Comment](./../../Comment/Comment 对象.html)** 对象，它代表与区域左上角单元格相关联的批注。

## 示例

```javascript
/*本示例给 A1 单元格添加批注 a1，显示批注的文本*/
function test() {
    Range("A1").AddComment("a1")
    console.log(Range("A1").Comment.Text())
}
```

```javascript
/*本示例给 A1 单元格添加批注 comment1，A2 单元格添加批注 comment2，显示 A1 单元格批注的下一条批注*/
function test() {
    Range("A1").AddComment("comment1")
    Range("A2").AddComment("comment2")
    console.log(Range("A1").Comment.Next().Text())
}
```
