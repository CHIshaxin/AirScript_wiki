---
title: "ProtectDrawingObjects (属性)"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/Worksheet/%E5%B1%9E%E6%80%A7/ProtectDrawingObjects%20%E5%B1%9E%E6%80%A7.html
updated: 1789700916000
hash: CYy5XDO-
---

# ProtectDrawingObjects (属性)

如果形状是受保护的，则为 **True**。要打开形状保护，请使用 **Protect** 方法，并将 *DrawingObjects* 参数设置为 **True**。**Boolean** 类型，只读。

## 示例

```javascript
/*本示例判断如果工作表 Sheet1 中的形状处于保护状态，则显示一个消息框。*/
function test() {
    if (Worksheets.Item("Sheet1").ProtectDrawingObjects) {
        console.log("The shapes on Sheet1 are protected.")
    }
}
```

```javascript
/*本示例判断工作簿中每一张工作表的形状是否处于保护状态。*/
function test() {
    for (let i = 1; i <= Application.Sheets.Count; i++) {
        let sheet = Application.Sheets.Item(i)
        if (sheet.ProtectDrawingObjects) {
            console.log(`工作表 ${sheet.Name} 中的形状处于保护状态`)
        } else {
            console.log(`工作表 ${sheet.Name} 中的形状未处于保护状态`)
        }
    }
}
```
