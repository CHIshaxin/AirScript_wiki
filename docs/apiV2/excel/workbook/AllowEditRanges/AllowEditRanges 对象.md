---
title: "AllowEditRanges 对象"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/AllowEditRanges/AllowEditRanges%20%E5%AF%B9%E8%B1%A1.html
---
# AllowEditRanges (对象)

所有 AllowEditRange对象的集合，这些对象代表受保护工作表上的可编辑单元格。

## 说明

使用 **[Protection](./../Protection/Protection 对象.html)** 对象的 **AllowEditRanges** 属性可返回 **[AllowEditRanges](./../AllowEditRanges/AllowEditRanges 对象.html)** 集合。

返回 **[AllowEditRanges](./../AllowEditRanges/AllowEditRanges 对象.html)** 集合后，可以使用 **Add** 方法添加可在受保护的工作表上编辑的区域。

## 示例

```javascript
/*在本示例中，ET 允许用户编辑活动工作表上的区域 A1:A4，并将指定区域的标题和地址通知用户。*/
function test() {
    let sheet = Application.ActiveSheet

    // Unprotect worksheet.
    sheet.Unprotect()
    wksPassword = "Enter password for the worksheet"

    // Establish a range that can allow edits on the protected worksheet.
    sheet.Protection.AllowEditRanges.Add("Classified", Range("A1:A4"), wksPassword)

    //Notify the user the title and address of the range.
    let allowEditRange = sheet.Protection.AllowEditRanges.Item(1)

    console.log(`Title of range: ${allowEditRange.Title}`)
    console.log(`Address of range: ${allowEditRange.Range.Address()}`)
}
```

```javascript
/*本示例显示第一张工作表上是否存在可编辑单元格。*/
function test() {
    console.log(Worksheets.Item(1).Protection.AllowEditRanges.Count > 0)
}
```
