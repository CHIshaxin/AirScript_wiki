---
title: "CheckSpelling (方法)"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/Application/%E6%96%B9%E6%B3%95/CheckSpelling%20%E6%96%B9%E6%B3%95.html
updated: 1789700916000
hash: B9FGAy2p
---

# CheckSpelling (方法)

对单个单词进行拼写检查。

## 说明

要检查工作表中的页眉、页脚和对象，可对 **[Worksheet](./../../Worksheet/Worksheet 对象.html)** 对象使用此方法。

要仅检查单元格和批注，可对 **Cells** 属性返回的对象使用此方法。

## 参数

| 属性 | 数据类型 | 必填 | 说明 |
| --- | --- | --- | --- |
| Word | string | 必填 | （仅应用于 **[Application](./../../Application/Application 对象.html)** 对象）要检查的单词。 |
| CustomDictionary | any | 可选 | 一个字符串，它表示自定义词典的文件名，如果在主词典中找不到单词，则会到此词典中查找。如果省略此参数，则使用当前指定的词典。 |
| IgnoreUppercase | any | 可选 | 如果为 **True**，则 ET 忽略所有字母都是大写的单词。如果为 **False**，则 ET 检查所有字母都是大写的单词。如果省略此参数，将使用当前的设置。 |

## 返回值

如果在任一词典中找到单词，则为 **True** ；否则为 **False**。

## 示例

```javascript
/*此示例演示CheckSpelling方法对单个单词拼写检查。*/
function test() {
    console.log("Hello是否拼写正确：" + Application.CheckSpelling("Hello"))
}
```

```javascript
/*此示例演示检查A1单元格上的大写单词KING。*/
function test() {
    Application.Range("A1").Value2 = "KING"
    console.log("检查结果：" + Application.CheckSpelling(Cells(1, 1).Value2, undefined, false))
}
```
