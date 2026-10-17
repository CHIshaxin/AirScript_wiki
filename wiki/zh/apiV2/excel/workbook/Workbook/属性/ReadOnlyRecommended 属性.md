---
title: "ReadOnlyRecommended (属性)"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/Workbook/%E5%B1%9E%E6%80%A7/ReadOnlyRecommended%20%E5%B1%9E%E6%80%A7.html
updated: 1789700916000
hash: DM6m60rs
---

# ReadOnlyRecommended (属性)

如果工作簿以建议只读方式保存，则该属性值为 **True**。**Boolean** 类型，只读。

## 说明

当打开一个建议为只读的工作簿时，ET 将显示一条信息，建议以只读方式打开此工作簿。

使用 **SaveAs** 方法可更改该属性。

## 示例

```javascript
/*本示例检查当前工作簿是否以建议只读方式保存，如果是，则显示一条消息。*/
function test() {
    if (ActiveWorkbook.ReadOnlyRecommended) {
        console.log("This workbook is saved as read-only recommended")
    }
}
```

```javascript
/*本示例演示如果活动工作簿不是以建议只读方式保存，则将其另存为建议只读，并告知用户。*/
function test() {
    if (!ActiveWorkbook.ReadOnlyRecommended) {
        ActiveWorkbook.SaveAs(ActiveWorkbook.FullName, undefined, undefined, undefined, true)
        if (ActiveWorkbook.ReadOnlyRecommended) {
            console.log("当前工作簿建议只读方式保存！")
        }
    }
}
```
