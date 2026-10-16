---
title: "Value (属性)"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/Borders/%E5%B1%9E%E6%80%A7/Value%20%E5%B1%9E%E6%80%A7.html
updated: 1789700916000
hash: Bl1DMXiL
---

# Value (属性)

等价于Borders.LineStyle

## 示例

```javascript
/*本示例将活动工作簿的第一个样式对象的边框线设置成 xlContinuous 。*/
function test() {
    ActiveWorkbook.Styles.Item(1).Borders.Value = xlContinuous
}
```

```javascript
/*本示例显示活动工作簿的第一个样式对象的边框线条样式是否为 xlNone 。*/
function test() {
    console.log(ActiveWorkbook.Styles.Item(1).Borders.Value == xlNone)
}
```
