---
title: "DisplayRightToLeft (属性)"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/ListObject/%E5%B1%9E%E6%80%A7/DisplayRightToLeft%20%E5%B1%9E%E6%80%A7.html
updated: 1789700916000
hash: CXAtPcLp
---

# DisplayRightToLeft (属性)

如果指定的 **ListObject**从右到左而不是从左到右显示，则此属性为 **True**。如果对象从左到右显示，则此属性为 **False**。**Boolean** 类型，只读。

## 示例

```javascript
/*本示例显示活动工作表上包含 B2 单元格的列表的是否从右到左展示。*/
function test() {
    let listObj = Application.ActiveSheet.Range("B2").ListObject
    console.log(listObj.DisplayRightToLeft)
}
```

```javascript
/*本示例判断工作表 Sheet1 上第一张列表的显示方式，并通知用户。*/
function test() {
    let listObj = Application.Worksheets.Item("Sheet1").ListObjects.Item(1)
    if (listObj.DisplayRightToLeft) {
        console.log("列表从右到左显示")
    } else {
        console.log("列表从左到右显示")
    }
}
```
