---
title: "ProtectContents 属性"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/Worksheet/%E5%B1%9E%E6%80%A7/ProtectContents%20%E5%B1%9E%E6%80%A7.html
---
# ProtectContents (属性)

如果工作表内容是受保护的，则为 **True**。此属性保护单独的单元格。要打开内容保护，请使用 **Protect** 方法，并将 *Contents* 参数设置为 **True**。**Boolean** 类型，只读。

## 示例

```javascript
/*本示例判断如果 sheet1 的内容处于保护状态，则显示一个消息框。*/
function test() {
    if (Worksheets.Item("Sheet1").ProtectContents) {
        console.log("The contents of Sheet1 are protected.")
    }
}
```

```javascript
/*本示例判断工作簿中每一张工作表的内容是否处于保护状态。*/
function test() {
    for (let i = 1; i <= Application.Sheets.Count; i++) {
        let sheet = Application.Sheets.Item(i)
        if (sheet.ProtectContents) {
            console.log(`工作表 ${sheet.Name} 的内容处于保护状态`)
        } else {
            console.log(`工作表 ${sheet.Name} 的内容未处于保护状态`)
        }
    }
}
```
