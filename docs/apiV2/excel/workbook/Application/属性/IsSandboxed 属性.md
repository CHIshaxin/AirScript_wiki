---
title: "IsSandboxed 属性"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/Application/%E5%B1%9E%E6%80%A7/IsSandboxed%20%E5%B1%9E%E6%80%A7.html
---
# IsSandboxed (属性)

如果指定的工作簿在受保护的视图窗口打开，则返回 **True**。只读。

## 说明

使用 **IsSandboxed** 属性可确定工作簿是否已在 **“受保护的视图”** 窗口中打开。

## 返回值

**Boolean**

## 示例

```javascript
/*本示例显示活动工作簿是否已在“受保护的视图”窗口中打开。*/
function test() {
    console.log(ActiveWorkbook.Application.IsSandboxed)
}
```

```javascript
/*本示例判断活动工作簿是否已在“受保护的视图”窗口中打开，并通知用户。*/
function test() {
    if (ActiveWorkbook.Application.IsSandboxed) {
        console.log("活动工作簿已在“受保护的视图”窗口中打开")
    } else {
        console.log("活动工作簿未在“受保护的视图”窗口中打开")
    }
}
```
