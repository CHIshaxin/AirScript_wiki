---
title: "ShrinkToFit 属性"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/Style/%E5%B1%9E%E6%80%A7/ShrinkToFit%20%E5%B1%9E%E6%80%A7.html
---
# ShrinkToFit (属性)

返回或设置一个 **Boolean** 值，它指明文本是否可以自动收缩为适当尺寸以适应可用列宽。

## 示例

```javascript
/*本示例显示活动工作簿第一个样式的文本是否可以自动适应列宽。*/
function test() {
    console.log(ActiveWorkbook.Styles.Item(1).ShrinkToFit)
}
```

```javascript
/*本示例为单元格 B1 赋值，并设置单元格的文本可以自动收缩为适当尺寸以适应可用列宽。*/
function test() {
    Range("B1").Value2 = "这是示例文本"
    Range("B1").Style.ShrinkToFit = true
}
```
