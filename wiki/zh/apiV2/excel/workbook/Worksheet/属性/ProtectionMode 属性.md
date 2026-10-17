---
title: "ProtectionMode (属性)"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/Worksheet/%E5%B1%9E%E6%80%A7/ProtectionMode%20%E5%B1%9E%E6%80%A7.html
updated: 1789700916000
hash: Dr60yNIf
---

# ProtectionMode (属性)

如果启用了用户界面专用保护，则为 **True**。要打开用户界面保护，请使用 **Protect** 方法，并将 *UserInterfaceOnly* 参数设置为 **True**。**Boolean** 类型，只读。

## 示例

```javascript
/*本示例显示活动工作表 ProtectionMode 属性的状态。*/
function test() {
    console.log(ActiveSheet.ProtectionMode)
}
```

```javascript
/*本示例判断工作簿中每一张工作表是否启用了用户界面专用保护。*/
function test() {
    for (let i = 1; i <= Sheets.Count; i++) {
        let sheet = Application.Sheets.Item(i)
        if (sheet.ProtectionMode) {
            console.log(`工作表 ${sheet.Name} 启用了用户界面专用保护`)
        } else {
            console.log(`工作表 ${sheet.Name} 未启用了用户界面专用保护`)
        }
    }
}
```
