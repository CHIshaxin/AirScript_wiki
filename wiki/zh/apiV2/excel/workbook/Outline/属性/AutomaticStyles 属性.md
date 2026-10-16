---
title: "AutomaticStyles (属性)"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/Outline/%E5%B1%9E%E6%80%A7/AutomaticStyles%20%E5%B1%9E%E6%80%A7.html
updated: 1789700916000
hash: B23FhKKN
---

# AutomaticStyles (属性)

如果分级显示使用自动样式，则该属性值为 **True**。**Boolean** 类型，可读写。

## 示例

```javascript
/*本示例将 Sheet1 的分级显示设置为使用自动样式。*/
function test() {
    Worksheets.Item("Sheet1").Outline.AutomaticStyles = true
}
```

```javascript
/*此示例显示活动工作表分级显示是否使用自动样式。*/
function test() {
    console.log(ActiveSheet.Outline.AutomaticStyles)
}
```
