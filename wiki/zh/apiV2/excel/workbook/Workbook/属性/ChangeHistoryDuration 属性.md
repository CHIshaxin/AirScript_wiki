---
title: "ChangeHistoryDuration (属性)"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/Workbook/%E5%B1%9E%E6%80%A7/ChangeHistoryDuration%20%E5%B1%9E%E6%80%A7.html
updated: 1789700916000
hash: GFYVZTd1
---

# ChangeHistoryDuration (属性)

返回或者设置共享工作簿的修订记录中所要显示修订信息的天数。**Long** 类型，可读写。

## 说明

关闭工作簿时，将删除修订记录中任何比该属性设置日期早的修订。

## 示例

```javascript
/*本示例用于设置追踪天数：当共享工作簿修订记录功能处于打开状态时，为活动工作簿的修订记录功能设置跟踪天数。关闭工作簿时，将删除修订记录中任何比该属性设置日期早的修订。*/
function test() {
    if (ActiveWorkbook.MultiUserEditing && ActiveWorkbook.KeepChangeHistory) {
        ActiveWorkbook.ChangeHistoryDuration = 7
    }
}
```

```javascript
/*本示例演示将活动工作簿另存为共享工作簿，当工作簿修订记录功能处于打开状态时，为活动工作簿的修订记录功能设置跟踪天数为五天。*/
function test() {
    ActiveWorkbook.SaveAs(ActiveWorkbook.FullName, undefined, undefined, undefined, undefined, undefined, xlShared)
    if (ActiveWorkbook.KeepChangeHistory) {
        ActiveWorkbook.ChangeHistoryDuration = 5
    }
}
```
