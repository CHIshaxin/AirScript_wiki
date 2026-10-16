---
title: "ProtectDrawingObjects 属性"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/Chart/%E5%B1%9E%E6%80%A7/ProtectDrawingObjects%20%E5%B1%9E%E6%80%A7.html
---
# ProtectDrawingObjects (属性)

如果形状是受保护的，则为 **True**。要打开形状保护，请使用 **Protect** 方法，并将 *DrawingObjects* 参数设置为 **True**。**Boolean** 类型，只读。

## 示例

```javascript
/*本示例显示图表工作表 Chart1 的形状是否受保护。*/
function test() {
    let chart = Application.Charts.Item("Chart1")
    console.log(chart.ProtectDrawingObjects)
}
```

```javascript
/*本示例判断图表工作表 Chart2 的形状是否受保护，并通知用户。*/
function test() {
    let chart = Application.Charts.Item("Chart2")
    if (chart.ProtectDrawingObjects) {
        console.log("工作表的形状受保护")
    } else {
        console.log("工作表的形状未受保护")
    }
}
```
