---
title: "NegativeBarFormat 属性"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/Databar/%E5%B1%9E%E6%80%A7/NegativeBarFormat%20%E5%B1%9E%E6%80%A7.html
---
# NegativeBarFormat (属性)

返回与数据条条件格式规则关联的 **NegativeBarFormat** 对象。只读。

## 返回值

**NegativeBarFormat**

## 示例

```javascript
/*本示例设置活动工作表上区域 A1:A10 中第一个条件格式（数据条）的负数据条的填充颜色。*/
function test() {
    let databar = ActiveSheet.Range("A1:A10").FormatConditions.Item(1)
    databar.NegativeBarFormat.Color.Color = RGB(255, 255, 0)
}
```

```javascript
/*本示例设置活动工作表上区域 A1:A10 中第一个条件格式（数据条）的负数据条的填充颜色和边框颜色。*/
function test() {
    let databar = ActiveSheet.Range("A1:A10").FormatConditions.Item(1)
    databar.NegativeBarFormat.Color.ColorIndex = 6
    databar.NegativeBarFormat.BorderColor.ColorIndex = 6
}
```
