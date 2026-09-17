---
title: "PasteSpecial 方法"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/Worksheet/%E6%96%B9%E6%B3%95/PasteSpecial%20%E6%96%B9%E6%B3%95.html
---
# PasteSpecial (方法)

以指定格式将剪贴板中的内容粘贴到工作表上。可用本方法从其他应用程序中粘贴数据，或以特定格式粘贴数据。

## 说明

仅当 *Format* =“HTML”时，*NoHTMLFormatting* 才起作用，所有其他情况下将忽略 *NoHTMLFormatting*。

必须在使用本方法之前选定目标区域。

该方法可能会修改工作表的选定区域，这取决于剪贴板中的内容。

## 参数

| 属性 | 数据类型 | 必填 | 说明 |
| --- | --- | --- | --- |
| Format | any | 可选 | 指定数据的剪贴板格式的字符串。 |
| Link | any | 可选 | 如果为 **True**，则链接到被粘贴数据的源。如果源数据不适于链接，或源应用程序不支持链接，将忽略此参数。默认值是 **False**。 |
| DisplayAsIcon | any | 可选 | 如果为 **True**，则将粘贴内容显示为图标。默认值是 **False**。 |
| IconFileName | any | 可选 | 如果 *DisplayAsIcon* 为 **True**，则指定包含所用图标的文件名。 |
| IconIndex | any | 可选 | 图标文件内的图标索引号。 |
| IconLabel | any | 可选 | 图标的文本标签。 |
| NoHTMLFormatting | any | 可选 | 如果为 **True**，则从 HTML 中删除所有的格式设置、超链接和图像。如果为 **False**，则完整粘贴 HTML。默认值是 **False**。 |

## 示例

```javascript
/*此示例将剪贴板中的“WPS文字 文件 对象”粘贴到工作表 Sheet1 上的单元格 D1 中。*/
function test() {
    Worksheets.Item("Sheet1").Range("D1").Select()
    Application.ActiveSheet.PasteSpecial("WPS文字 文件 对象")
}
```

```javascript
/*此示例在工作表 Sheet1 复制单元格A1中的图片对象粘贴到单元格 F1 中。*/
function test() {
    Worksheets.Item("Sheet1").Range("A1").Copy()
    Worksheets.Item("Sheet1").Range("F1").Select()
    ActiveSheet.PasteSpecial(0)
}
```
