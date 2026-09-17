---
title: "ProtectStructure 属性"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/Workbook/%E5%B1%9E%E6%80%A7/ProtectStructure%20%E5%B1%9E%E6%80%A7.html
---
# ProtectStructure (属性)

如果工作簿中的工作表次序处于保护状态，则该属性值为 **True**。**Boolean** 类型，只读。

## 示例

```javascript
/*本示例检查活动工作簿中工作表结构是否受保护，如果是则显示一条消息。*/
function test() {
    if (ActiveWorkbook.ProtectStructure) {
        console.log("Remember, you cannot delete, add, or change the location of any sheets in this workbook.")
    }
}
```

```javascript
/*本示例演示第一个工作簿先开启保护工作簿，再检查工作表结构是否受保护，若受保护则提示已处于保护状态。*/
function test() {
    Application.Workbooks.Item(1).Protect("Y6dh!et5", true, true)
    if (Workbooks.Item(1).ProtectStructure) {
        console.log("工作表次序已处于保护状态")
    }
}
```
