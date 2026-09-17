---
title: "ProtectionMode 属性"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/Chart/%E5%B1%9E%E6%80%A7/ProtectionMode%20%E5%B1%9E%E6%80%A7.html
---
# ProtectionMode (属性)

如果启用了用户界面专用保护，则为 **True**。要打开用户界面保护，请使用 **Protect** 方法，并将 *UserInterfaceOnly* 参数设置为 **True**。**Boolean** 类型，只读。

## 示例

```javascript
/*此示例显示活动工作表 ProtectionMode 属性的状态。*/
function test() {
    console.log(Application.ActiveSheet.ProtectionMode)
}
```

```javascript
/*本示例保护图表工作表 Chart1，然后判断是否启用用户界面并通知用户。*/
function test() {
    let chart = Application.Charts.Item("Chart1")
    chart.Protect(null, true, false, true, true)
    if (chart.ProtectionMode == false) {
        console.log("工作表的用户界面未受保护")
    } else {
        console.log("工作表的用户界面受保护")
    }
}
```
