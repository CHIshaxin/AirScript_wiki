---
title: "ProtectContents 属性"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/Chart/%E5%B1%9E%E6%80%A7/ProtectContents%20%E5%B1%9E%E6%80%A7.html
---
# ProtectContents (属性)

如果工作表内容是受保护的，则为 **True**。对于图表，这样会保护整个图表。要打开内容保护，请使用 **Protect** 方法，并将 *Contents* 参数设置为 **True**。**Boolean** 类型，只读。

## 示例

```javascript
/*本示例显示图表工作表 Chart1 的内容是否受到保护。*/
function test() {
    let chart = Application.Charts.Item("Chart1")
    console.log(chart.ProtectContents)
}
```

```javascript
/*本示例判断图表工作表 Chart2 的内容是否受保护，并通知用户。*/
function test() {
    let chart = Application.Charts.Item("Chart2")
    if (chart.ProtectContents) {
        console.log("工作表的内容受保护")
    } else {
        console.log("工作表的内容未受保护")
    }
}
```
