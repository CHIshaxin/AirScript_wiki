---
title: "ProtectWindows 属性"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/Workbook/%E5%B1%9E%E6%80%A7/ProtectWindows%20%E5%B1%9E%E6%80%A7.html
---
# ProtectWindows (属性)

如果工作簿窗口受保护，则该属性值为 **True**。**Boolean** 类型，只读。

## 示例

```javascript
/*本示例检查活动工作簿窗口是否受保护，如果是则显示一条消息。*/
function test() {
    if (ActiveWorkbook.ProtectWindows) {
        console.log("Remember, you cannot rearrange any window in this workbook.")
    }
}
```

```javascript
/*本示例演示活动工作簿先开启保护工作簿窗口，再检查活动工作簿窗口是否受保护，若受保护则提示已处于保护状态。*/
function test() {
    ActiveWorkbook.Protect(undefined, undefined, true)
    if (ActiveWorkbook.ProtectWindows) {
        console.log("活动工作簿窗口已处于保护状态。")
    }
}
```
