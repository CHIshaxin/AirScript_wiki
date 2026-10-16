---
title: "ShowValue (属性)"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/Databar/%E5%B1%9E%E6%80%A7/ShowValue%20%E5%B1%9E%E6%80%A7.html
updated: 1789700916000
hash: h2YNEbn7
---

# ShowValue (属性)

返回或设置一个 **Boolean** 值，该值指定数据条条件格式应用于区域时是否显示单元格中的值。

## 示例

```javascript
/*本示例显示活动工作表上区域 A1:A10 中第一个条件格式（数据条）是否显示单元格中的值。*/
function test() {
    let databar = ActiveSheet.Range("A1:A10").FormatConditions.Item(1)
    console.log(databar.ShowValue)
}
```

```javascript
/*本示例将活动工作表上区域 A1:A10 中第一个条件格式（数据条）设置为不展示单元格中的值，并将该数据条填充色设置为纯色。*/
function test() {
    let databar = ActiveSheet.Range("A1:A10").FormatConditions.Item(1)
    databar.ShowValue = false
    databar.BarFillType = xlDataBarFillSolid
}
```
