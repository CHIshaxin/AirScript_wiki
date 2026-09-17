---
title: "Paste 方法"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/Worksheet/%E6%96%B9%E6%B3%95/Paste%20%E6%96%B9%E6%B3%95.html
---
# Paste (方法)

将“剪贴板”中的内容粘贴到工作表上。

## 说明

如果不指定 *Destination* 参数，则在使用该方法之前必须选择目标区域。

该方法可能会修改工作表的选定区域，这取决于剪贴板中的内容。

## 参数

| 属性 | 数据类型 | 必填 | 说明 |
| --- | --- | --- | --- |
| Destination | any | 可选 | 一个 **[Range](./../../Range/Range 对象.html)** 对象，指定用于粘贴剪贴板中内容的目标区域。如果省略此参数，就使用当前的选定区域。仅当剪贴板中的内容能被粘贴到某区域时，才能指定此参数。如果指定了此参数，则不能使用 *Link* 参数。 |
| Link | any | 可选 | 如果为 **True**，则链接到被粘贴数据的源。如果指定此参数，则不能使用 *Destination* 参数。默认值是 **False**。 |

## 示例

```javascript
/*此示例将工作表 Sheet1 上单元格区域 C1:C5 中的数据复制到单元格区域 D1:D5 中。*/
function test() {
    Worksheets.Item("Sheet1").Range("C1:C5").Copy()
    Worksheets.Item("Sheet1").Paste(Worksheets.Item("Sheet1").Range("D1:D5"))
}
```

```javascript
/*本示例复制 A1 单元格后，选中 B1 单元格再粘贴，并指定链接到被粘贴数据的源。*/
function test() {
    ActiveSheet.Range("A1").Copy()
    ActiveSheet.Range("B1").Select()
    ActiveSheet.Paste(null, true)
}
```
