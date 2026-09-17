---
title: "Next 方法"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/Comment/%E6%96%B9%E6%B3%95/Next%20%E6%96%B9%E6%B3%95.html
---
# Next (方法)

返回一个 **[Comment](./../../Comment/Comment 对象.html)** 对象，该对象代表下一条批注。

## 说明

本方法仅对单张工作表有效。对工作表中最后一条批注使用本方法可返回 **Null**（不是下一张工作表的第一条批注）。

在没有现有注释的新工作簿中测试此示例。 若要清除工作簿中的所有批注，请在“即时”窗格中使用 Selection.SpecialCells(xlCellTypeComments).Delete() 。

## 返回值

Comment

## 示例

```javascript
/*本示例隐藏下一条批注。*/
function test() {
    //Sets up the comments
    for (let xNum = 1; xNum <= 10; xNum++) {
        Application.Range("A" + xNum).AddComment()
        Application.Range("A" + xNum).Comment.Text("Comment " + xNum)
    }
    console.log("Comments created... A1:A10")
    //Deletes every second comment in the A1:A10 range
    for (let yNum = 1; yNum <= 10; yNum = yNum + 2) {
        Application.Range("A" + yNum).Comment.Next().Shape.Select(true)
        Application.Selection.Delete()
    }
    console.log("Deleted every second comment")
}
```

```javascript
/*本示例显示 A3 单元格下一条批注的内容。*/
function test() {
    console.log(Application.Range("A3").Comment.Next().Text())
}
```
