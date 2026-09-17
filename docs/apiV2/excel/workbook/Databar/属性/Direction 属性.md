---
title: "Direction 属性"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/Databar/%E5%B1%9E%E6%80%A7/Direction%20%E5%B1%9E%E6%80%A7.html
---
# Direction (属性)

返回或设置数据条的显示方向。可读/写。

## 说明

可以是以下 **XlReadingOrder** 之一：**xlContext**、**xlLTR**（左至右）或**xlRTL**（右至左）。

## 示例

```javascript
/*本示例显示活动工作表上区域 A1:A10 中第一个为数据条的条件格式的展示方向是否为xlContext。*/
function test() {
    let databar = ActiveSheet.Range("A1:A10").FormatConditions.Item(1)
    console.log(databar.Direction == xlContext)
}
```

```javascript
/*本示例设置活动工作表上区域 A1:A10 上第一个为数据条的条件格式的展示方向。*/
function test() {
    let databar = ActiveSheet.Range("A1:A10").FormatConditions.Item(1)
    databar.Direction = xlRTL
}
```
