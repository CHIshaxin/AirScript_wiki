---
title: "PhoneticCharacters 属性"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/Characters/%E5%B1%9E%E6%80%A7/PhoneticCharacters%20%E5%B1%9E%E6%80%A7.html
---
# PhoneticCharacters (属性)

返回或设置指定 **[Characters](./../../Characters/Characters 对象.html)** 对象中的拼音文本。String 类型，可读写。

## 说明

要在单元格中添加拼音信息，请不要使用该属性，应使用 **Phonetics** 集合的 **Add** 方法；另外，应使用 **Phonetic** 对象的 **Text** 属性返回或设置单元格中的拼音文本字符串。

## 示例

```javascript
/*本示例将活动单元格中从文本起始位置开始的第四个字符替换。*/
ActiveCell.Characters(1,3).PhoneticCharacters = "替换字符"
```
