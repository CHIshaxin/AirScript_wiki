---
title: "Previous (方法)"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/Comment/%E6%96%B9%E6%B3%95/Previous%20%E6%96%B9%E6%B3%95.html
updated: 1789700916000
hash: FyJOxbwA
---

# Previous (方法)

返回一个 **[Comment](./../../Comment/Comment 对象.html)** 对象，该对象代表前一条批注。

## 说明

本方法仅对单张工作表有效。对工作表中第一条批注使用本方法可返回 **Null**（不是前一张工作表的最后一条批注）。

在没有现有注释的新工作簿中测试此示例。 若要清除工作簿中的所有批注，请在“即时”窗格中使用 Selection.SpecialCells(xlCellTypeComments).Delete() 。

## 返回值

Comment

## 示例

```javascript
/*本示例使用 Previous 方法导航，每隔一条删除评论。*/
function test() {
    //Sets up the comments
    for (let xNum = 1; xNum <= 10; xNum++) {
        Application.Range("A" + xNum).AddComment()
        Application.Range("A" + xNum).Comment.Text("Comment " + xNum)
    }
    console.log("Comments created... A1:A10")
    //Deletes every second comment in the A1:A10 range
    for (let yNum = 10; yNum >= 1; yNum = yNum - 2) {
        Application.Range("A" + yNum).Comment.Previous().Shape.Select(true)
        Application.Selection.Delete()
    }
    console.log("Deleted every second comment")
}
```

```javascript
/*本示例显示 A4 单元格上一条批注的内容。*/
function test() {
    console.log(Application.Range("A4").Comment.Previous().Text())
}
```
