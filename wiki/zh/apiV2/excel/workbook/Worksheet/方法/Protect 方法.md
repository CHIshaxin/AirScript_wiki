---
title: "Protect (方法)"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/Worksheet/%E6%96%B9%E6%B3%95/Protect%20%E6%96%B9%E6%B3%95.html
updated: 1789700916000
hash: zqFwvN9C
---

# Protect (方法)

保护工作表使其不能被修改。

## 说明

要在受保护的工作表上做更改，如果提供密码，则可在受保护的工作表上使用 **Protect** 方法。另一种方法是：取消工作表保护，对工作表做一些必要的更改，然后再次保护工作表。

**注释：**“取消保护”的意思是单元格可以被锁定（“设置单元格格式”对话框），但在“允许用户编辑区域”对话框中定义的单元格区域内，并且用户已通过密码取消了对单元格区域的保护或已通过 NT 权限的验证。

## 参数

| 属性 | 数据类型 | 必填 | 说明 |
| --- | --- | --- | --- |
| Password | any | 可选 | 一个字符串，该字符串为工作表或工作簿指定区分大小写的密码。如果省略此参数，不用密码就可以取消对工作表或工作簿的保护。否则，必须指定密码才能取消对工作表或工作簿的保护。如果忘记了密码，就无法取消对工作表或工作簿的保护。 |
| DrawingObjects | any | 可选 | 如果为 True，则保护形状。默认值是 True。 |
| Contents | any | 可选 | 如果为 True，则保护内容。对于图表，这样会保护整个图表。对于工作表，这样会保护锁定的单元格。默认值是 True。 |
| Scenarios | any | 可选 | 如果为 True，则保护方案。此参数仅对工作表有效。默认值是 True。 |
| UserInterfaceOnly | any | 可选 | 如果为 True，则保护用户界面，但不保护宏。如果省略此参数，则既保护宏也保护用户界面。 |
| AllowFormattingCells | any | 可选 | 如果为 True，则允许用户为受保护的工作表上的任意单元格设置格式。默认值是 False。 |
| AllowFormattingColumns | any | 可选 | 如果为 True，则允许用户为受保护的工作表上的任意列设置格式。默认值是 False。 |
| AllowFormattingRows | any | 可选 | 如果为 True，则允许用户为受保护的工作表上的任意行设置格式。默认值是 False。 |
| AllowInsertingColumns | any | 可选 | 如果为 True，则允许用户在受保护的工作表上插入列。默认值是 False。 |
| AllowInsertingRows | any | 可选 | 如果为 True，则允许用户在受保护的工作表上插入行。默认值是 False。 |
| AllowInsertingHyperlinks | any | 可选 | 如果为 True，则允许用户在受保护的工作表中插入超链接。默认值是 False。 |
| AllowDeletingColumns | any | 可选 | 如果为 True，则允许用户在受保护的工作表上删除列，要删除的列中的每个单元格都被解除锁定。默认值是 False。 |
| AllowDeletingRows | any | 可选 | 如果为 True，则允许用户在受保护的工作表上删除行，要删除的行中的每个单元格都被解除锁定。默认值是 False。 |
| AllowSorting | any | 可选 | 如果为 True，则允许用户在受保护的工作表上进行排序。排序区域中的每个单元格必须是解除锁定的或取消保护的。默认值是 False。 |
| AllowFiltering | any | 可选 | 如果为 True，则允许用户在受保护的工作表上设置筛选。用户可以更改筛选条件，但是不能启用或禁用自动筛选功能。用户也可以在已有的自动筛选功能上设置筛选。默认值是 False。 |
| AllowUsingPivotTables | any | 可选 | 如果为 True，则允许用户在受保护的工作表上使用数据透视表。默认值是 False。 |

## 示例

```javascript
/*本示例对活动工作表设置密码进行保护，然后将此状态通知用户。*/
function test() {
    Application.ActiveSheet.Protect("Y6dh!et5", true, false, false, true)
    console.log("活动工作表已设置密码保护.")
}
```

```javascript
/*本示例判断如果第一张工作表内容未开启保护,则设置密码对第一张工作表形状及内容进行保护。*/
function test() {
    if (Application.Sheets.Item(1).ProtectContents == false) {
        Application.Sheets.Item(1).Protect("Y6dh!et5", true, true, false)
    }
}
```
