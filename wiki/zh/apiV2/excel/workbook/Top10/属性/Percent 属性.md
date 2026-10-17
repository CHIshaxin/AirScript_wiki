---
title: "Percent (属性)"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/Top10/%E5%B1%9E%E6%80%A7/Percent%20%E5%B1%9E%E6%80%A7.html
updated: 1789700916000
hash: Coqhbcrr
---

# Percent (属性)

返回或设置一个 **Boolean** 值，指定是否按百分比值确定排位。

## 示例

```javascript
/*本示例显示第一张工作表上单元格区域 C1:C10 的第二个（Top10）条件格式是否按百分比值确定排位。*/
function test() {
    let top = Application.Worksheets.Item(1).Range("C1:C10").FormatConditions.Item(2)
    console.log(top.Percent)
}
```

```javascript
/*本示例将第一张工作表上单元格区域 C1:C10 的第一个（Top10）条件格式设置为按百分比值确定排位。*/
function test() {
    let top = Application.Worksheets.Item(1).Range("C1:C10").FormatConditions.Item(1)
    top.Percent = true
}
```
