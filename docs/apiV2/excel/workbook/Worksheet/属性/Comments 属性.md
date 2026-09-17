---
title: "Comments 属性"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/Worksheet/%E5%B1%9E%E6%80%A7/Comments%20%E5%B1%9E%E6%80%A7.html
---
# Comments (属性)

返回一个 **[Comments](./../../Comments/Comments 对象.html)** 集合，该集合表示指定工作表的所有注释。只读。

## 示例

```javascript
/*本示例删除活动工作表中由 Jean Selva 添加的所有批注。*/
function test() {
    for (let i = 1; i <= ActiveSheet.Comments.Count; i++) {
        let c = ActiveSheet.Comments.Item(i)
        if (c.Author == "Jean Selva") {
            c.Delete()
        }
    }
}
```

```javascript
/*本示例在第一张工作表A1单元格添加一条批注，并将所有批注设置为隐藏，并通知用户批注数量。*/
function test() {
    Worksheets.Item(1).Range("A1").AddComment("This is a tips")
    Worksheets.Item(1).Comments.Visible = false
    console.log(Worksheets.Item(1).Comments.Count)
}
```
