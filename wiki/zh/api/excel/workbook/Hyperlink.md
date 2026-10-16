---
title: "Hyperlink"
source: https://airsheet.wps.cn/docs/api/excel/workbook/Hyperlink.html
updated: 1789700916000
hash: C2E-55pE
---

# Hyperlink

单个超链接对象，有转跳地址和显示文本两个属性，两者可以不相等。

Hyperlink 对象的具体属性和方法请参阅下方的列表。

### 属性列表

| 属性名 | 数据类型 | 简介 |
|---|---|---|
| [Address](#address) | String | 超链接转跳的地址 |
| [TextToDisplay](#texttodisplay) | String | 超链接显示的文本 |

## Address

设置/获取执行超链接的地址

### 数据类型

String - 超链接转跳的地址

### 示例

```js
//获取超链接对象
const hyperlink = Application.ActiveSheet.Hyperlinks.Item(1)
// 打印超链接的地址
console.log(hyperlink.Address) // https://www.kdocs.cn
hyperlink.Address = "https://www.kdocs.cn"
```

## TextToDisplay

设置/获取执行超链接的文本

### 数据类型

String - 超链接显示的文本

### 示例

```js
//获取超链接对象
const hyperlink = Application.ActiveSheet.Hyperlinks.Item(1)
// 打印超链接的地址
console.log(hyperlink.TextToDisplay) // 金山文档官网
hyperlink.TextToDisplay = "金山文档"
```
