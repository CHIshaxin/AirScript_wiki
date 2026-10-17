---
title: "ProtectScenarios (属性)"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/Worksheet/%E5%B1%9E%E6%80%A7/ProtectScenarios%20%E5%B1%9E%E6%80%A7.html
updated: 1789700916000
hash: DW1gtAzQ
---

# ProtectScenarios (属性)

如果工作表的方案处于保护状态，则该属性值为 **True**。**Boolean** 类型，只读。

## 示例

```javascript
/*本示例判断如果 sheet1 中的方案处于保护状态，则将显示消息框。*/
function test() {
    if (Worksheets.Item("Sheet1").ProtectScenarios) {
        console.log("Scenarios are protected on this worksheet.")
    }
}
```

```javascript
/*本示例判断工作簿中每一张工作表的方案是否处于保护状态。*/
function test() {
    for (let i = 1; i <= Sheets.Count; i++) {
        let sheet = Application.Sheets.Item(i)
        if (sheet.ProtectScenarios) {
            console.log(`工作表 ${sheet.Name} 方案处于保护状态`)
        } else {
            console.log(`工作表 ${sheet.Name} 方案未处于保护状态`)
        }
    }
}
```
