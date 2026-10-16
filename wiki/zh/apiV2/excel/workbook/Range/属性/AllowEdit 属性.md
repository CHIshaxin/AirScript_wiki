---
title: "AllowEdit (属性)"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/Range/%E5%B1%9E%E6%80%A7/AllowEdit%20%E5%B1%9E%E6%80%A7.html
updated: 1789700916000
hash: Bwo8dX-0
---

# AllowEdit (属性)

返回一个 **Boolean** 值，它指明是否可以在受保护的工作表上编辑区域。

## 示例

```javascript
/*本示例通知用户是否可编辑受保护的工作表上的单元格 A1。*/
function test() {
    let wksOne = Application.ActiveSheet
    //Protect the worksheet
    wksOne.Protect()
    //Notify the user about editing cell A1.
    if (wksOne.Range("A1").AllowEdit) {
        console.log("Cell A1 can be edited.");
    } else {
        console.log("Cell A1 cannot be edited.");
    }
}
```

```javascript
/*本示例新增一个工作表，然后将第二张工作表设为保护工作表，显示它们的可编辑状态*/
function test() {
    Worksheets.Add()
    let ws1 = Worksheets.Item(1)
    let ws2 = Worksheets.Item(2)
    ws2.Protect()
    if (ws1.Cells.AllowEdit != ws2.Cells.AllowEdit) {
        console.log("当前两个工作表只有一个允许编辑")
    }
}
```
