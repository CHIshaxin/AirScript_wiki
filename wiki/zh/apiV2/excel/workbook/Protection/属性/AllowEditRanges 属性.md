---
title: "AllowEditRanges (属性)"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/Protection/%E5%B1%9E%E6%80%A7/AllowEditRanges%20%E5%B1%9E%E6%80%A7.html
updated: 1789700916000
hash: Czz5v_am
---

# AllowEditRanges (属性)

返回 **[AllowEditRanges](./../../AllowEditRanges/AllowEditRanges 对象.html)** 对象。

## 示例

```javascript
/*在本示例中，ET 允许用户编辑活动工作表上的区域 A1:A4，并将指定区域的标题和地址通知用户。*/
function test() {
    let sheet = Application.ActiveSheet

    // Unprotect worksheet.
    sheet.Unprotect()

    // Establish a range that can allow edits on the protected worksheet.
    sheet.Protection.AllowEditRanges.Add("Classified", Range("A1:A4"), "123")

    //Notify the user the title and address of the range.
    let allowEditRange = sheet.Protection.AllowEditRanges.Item(1)

    console.log(`Title of range: ${allowEditRange.Title}`)
    console.log(`Address of range: ${allowEditRange.Range.Address()}`)
}
```

```javascript
/*本示例显示第二张工作表上 AllowEditRange 对象的数量。*/
function test() {
    console.log(Worksheets.Item(2).Protection.AllowEditRanges.Count)
}
```
