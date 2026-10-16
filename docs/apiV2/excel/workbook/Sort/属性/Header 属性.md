---
title: "Header 属性"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/Sort/%E5%B1%9E%E6%80%A7/Header%20%E5%B1%9E%E6%80%A7.html
---
# Header (属性)

指定第一行是否包含标题信息。可读/写 **[XlYesNoGuess](./../../枚举/XlYesNoGuess 枚举.html)** 类型。

## 说明

默认值为 **xlNo**。如果希望 ET 确定标题，可以指定 **xlGuess**。

## 示例

```javascript
/*本示例将活动工作表设置为由 ET 确定排序时是否包含标题。*/
function test() {
    Application.ActiveSheet.Sort.Header = xlGuess
}
```

```javascript
/*本示例设置第一张工作表排序时第一行包含标题信息。*/
function test() {
    Application.Sheets.Item(1).Sort.Header = xlYes
}
```
