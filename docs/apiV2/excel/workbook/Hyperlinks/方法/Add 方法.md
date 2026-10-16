---
title: "Add 方法"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/Hyperlinks/%E6%96%B9%E6%B3%95/Add%20%E6%96%B9%E6%B3%95.html
---
# Add (方法)

向指定的区域或形状添加超链接。返回一个 **[Hyperlink](./../../Hyperlink/Hyperlink 对象.html)** 对象，它代表新的超链接。

## 说明

指定 **TextToDisplay** 参数时，文本必须是字符串。

## 参数

| 属性 | 数据类型 | 必填 | 说明 |
| --- | --- | --- | --- |
| Anchor | IDispatch | 必填 | 超链接的位置。可为 Range 或 Shape 对象。 |
| Address | string | 必填 | 超链接的地址。 |
| SubAddress | any | 可选 | 超链接的子地址。 |
| ScreenTip | any | 可选 | 当鼠标指针停留在超链接上时所显示的屏幕提示。 |
| TextToDisplay | any | 可选 | 要显示的超链接的文本。 |

## 示例

```javascript
/*本示例向单元格 A5 添加超链接。*/
function test() {
    let sheet1 = Application.Worksheets.Item(1)
    sheet1.Hyperlinks.Add(sheet1.Range("A5"), "http://www.kingsoft.com", null, "kingsoft Web Site", "kingsoft")
}
```

```javascript
/*本示例向单元格 B5 添加一个电子邮件超链接。*/
function test() {
    let sheet1 = Application.Worksheets.Item(1)
    sheet1.Hyperlinks.Add(sheet1.Range("B5"), "mailto:someone@example.com?subject=hello", null, "Write us today", "Support")
}
```
