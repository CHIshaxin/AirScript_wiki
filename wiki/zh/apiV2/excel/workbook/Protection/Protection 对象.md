---
title: "Protection (对象)"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/Protection/Protection%20%E5%AF%B9%E8%B1%A1.html
updated: 1789700916000
hash: Cpb2Rk5t
---

# Protection (对象)

代表工作表可使用的各种保护选项类型。

## 说明

使用 **[Worksheet](./../Worksheet/Worksheet 对象.html)** 对象的 **Protection** 属性可返回一个 **[Protection](./../Protection/Protection 对象.html)** 对象。

返回一个 **[Protection](./../Protection/Protection 对象.html)** 对象后，就可用该对象的下列属性来设置或返回保护选项。

- **AllowDeletingColumns**
- **AllowDeletingRows**
- **AllowFiltering**
- **AllowFormattingCells**
- **AllowFormattingColumns**
- **AllowFormattingRows**
- **AllowInsertingColumns**
- **AllowInsertingHyperlinks**
- **AllowInsertingRows**
- **AllowSorting**
- **AllowUsingPivotTables**

## 示例

```javascript
/*下例通过在最上面的行中放三个成员并保留该工作表说明了如何使用 Protection 对象的 AllowInsertingColumns 属性。然后，此示例检查允许插入列的保护设置是否为 false ，并在必要时将其设置为 true。最后，通知用户插入一个列。*/
function test() {
    Range("A1").Formula = "1"
    Range("B1").Formula = "3"
    Range("C1").Formula = "4"
    ActiveSheet.Protect()
    // Check the protection setting of the worksheet and act accordingly.
    if (Application.ActiveSheet.Protection.AllowInsertingColumns == false) {
        Application.ActiveSheet.Protect(null, null, null, null, null, null, null, null, true)
        console.log("Insert a column between 1 and 3")
    } else {
        console.log("Insert a column between 1 and 3")
    }
}
```

```javascript
/*本示例显示是否允许在受保护的第二张工作表上插入列。*/
function test() {
    console.log(Worksheets.Item(2).Protection.AllowInsertingColumns)
}
```
