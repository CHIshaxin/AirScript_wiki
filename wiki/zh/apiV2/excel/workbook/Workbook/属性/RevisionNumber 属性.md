---
title: "RevisionNumber (属性)"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/Workbook/%E5%B1%9E%E6%80%A7/RevisionNumber%20%E5%B1%9E%E6%80%A7.html
updated: 1789700916000
hash: es73Xw6u
---

# RevisionNumber (属性)

返回指定工作簿以共享清单方式打开后已保存的次数。如果该工作簿以独占模式打开，则该值为 0（零）。**Long** 类型，只读。

## 说明

仅在保存工作簿的本地副本时，**RevisionNumber** 属性才进行更新，保存该工作簿的远程副本时该属性不进行更新。

## 示例

```javascript
/*本示例使用修订号判断活动工作簿是否以独占模式打开。如果是则本示例以共享清单方式保存该工作簿。*/
function test() {
    if (ActiveWorkbook.RevisionNumber == 0) {
        ActiveWorkbook.SaveAs(ActiveWorkbook.FullName, undefined, undefined, undefined, undefined, undefined, xlShared, xlOtherSessionChanges)
    }
}
```

```javascript
/*本示例演示RevisionNumber属性的数据类型。*/
function test() {
    console.log(typeof ActiveWorkbook.RevisionNumber)
}
```
