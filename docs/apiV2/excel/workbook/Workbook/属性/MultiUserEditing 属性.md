---
title: "MultiUserEditing 属性"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/Workbook/%E5%B1%9E%E6%80%A7/MultiUserEditing%20%E5%B1%9E%E6%80%A7.html
---
# MultiUserEditing (属性)

如果工作簿以共享列表方式打开，则该属性值为 **True**。**Boolean** 类型，只读。

## 说明

若要以共享清单模式保存工作簿，请使用 **SaveAs** 方法。若要将工作簿访问方式从共享模式切换到独占模式，请使用 **ExclusiveAccess** 方法。

## 示例

```javascript
/*本示例检查活动工作簿是否以独占模式打开。如果是则本示例以共享清单方式保存该工作簿。*/
function test() {
    if (!ActiveWorkbook.MultiUserEditing) {
        ActiveWorkbook.SaveAs(ActiveWorkbook.FullName, undefined, undefined, undefined, undefined, undefined, xlShared)
    }
}
```

```javascript
/*本示例演示将第一个工作簿设为共享，告知用户后重新设为独占访问，并告知用户。*/
function test() {
    let wb = Workbooks.Item(1)
    wb.SaveAs(wb.FullName, undefined, undefined, undefined, undefined, undefined, xlShared)
    console.log("是否为共享工作簿：" + wb.MultiUserEditing)
    wb.ExclusiveAccess()
    console.log(wb.Name + "已改为独占访问。")
}
```
