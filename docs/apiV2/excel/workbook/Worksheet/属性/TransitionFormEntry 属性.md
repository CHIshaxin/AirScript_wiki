---
title: "TransitionFormEntry 属性"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/Worksheet/%E5%B1%9E%E6%80%A7/TransitionFormEntry%20%E5%B1%9E%E6%80%A7.html
---
# TransitionFormEntry (属性)

如果 ET 对工作表使用 Lotus 1-2-3 的公式输入规则，则该属性的值为 **True**。**Boolean** 类型，可读写。

## 示例

```javascript
/*此示例对 Sheet1 使用 Lotus 1-2-3 公式输入规则。*/
function test() {
    Worksheets.Item("Sheet1").TransitionFormEntry = true
}
```

```javascript
/*此示例判断如果活动工作表未使用 Lotus 1-2-3 公式输入规则，则将活动工作表设置为使用 Lotus 1-2-3 公式输入规则。*/
function test() {
    if (Application.ActiveSheet.TransitionFormEntry == false) {
        Application.ActiveSheet.TransitionFormEntry = true
        console.log("已将活动工作表设置为使用 Lotus 1-2-3 公式输入规则")
    } else {
        console.log("活动工作表使用 Lotus 1-2-3 公式输入规则")
    }
}
```
