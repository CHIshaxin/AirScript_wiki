---
title: "EmailSubject (属性)"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/Hyperlink/%E5%B1%9E%E6%80%A7/EmailSubject%20%E5%B1%9E%E6%80%A7.html
updated: 1789700916000
hash: SqmFf4cK
---

# EmailSubject (属性)

返回或设置指定超链接的电子邮件主题行的文本字符串。主题行是添加到超链接地址上的。**String** 类型，可读写。

## 说明

该属性常与电子邮件超链接一起使用。

该属性的值优先于使用同一 **[Hyperlink](./../../Hyperlink/Hyperlink 对象.html)** 对象的 **Address** 属性指定的任何电子邮件主题行。

## 示例

```javascript
/*本示例设置第一张工作表中第一个超链接的电子邮件主题行*/
function test() {
    Application.Worksheets.Item(1).Hyperlinks.Item(1).EmailSubject = "Quote Request"
}
```

```javascript
/*本示例设置单元格G1中的第一项超链接的电子邮件主题行*/
function test() {
    let hl = Range("G1").Hyperlinks.Item(1)
    hl.EmailSubject = "活动报名"
}
```
