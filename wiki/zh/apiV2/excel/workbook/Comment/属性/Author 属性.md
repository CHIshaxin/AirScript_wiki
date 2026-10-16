---
title: "Author (属性)"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/Comment/%E5%B1%9E%E6%80%A7/Author%20%E5%B1%9E%E6%80%A7.html
updated: 1789700916000
hash: D2u-XO0V
---

# Author (属性)

返回或设置批注的作者。**String** 类型，只读。

## 示例

```javascript
/*本示例删除活动工作表中由 Jean Selva 添加的所有批注。*/
function test() {
    let myComments = Application.ActiveSheet.Comments
    for (let c = 1; c <= myComments.Count; c++) {
        if (myComments.Item(c).Author == "Jean Selva") {
            myComments.Item(c).Delete()
        }
    }
}
```

```javascript
/*本示例显示第一张工作表中 D5 单元格中批注的作者。*/
function test() {
    console.log(Worksheets.Item(1).Range("D5").Comment.Author)
}
```
