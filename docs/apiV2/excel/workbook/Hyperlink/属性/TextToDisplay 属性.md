---
title: "TextToDisplay 属性"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/Hyperlink/%E5%B1%9E%E6%80%A7/TextToDisplay%20%E5%B1%9E%E6%80%A7.html
---
# TextToDisplay (属性)

返回或设置要为指定超链接显示的文本。默认值为超链接的地址。**String** 类型，可读写。

## 示例

```javascript
/*本示例为活动工作表上第一个超链接设置要显示的文本*/
function test() {
    Application.ActiveSheet.Hyperlinks.Item(1).TextToDisplay = "Company Home Page"
}
```

```javascript
/*本示例修改B5单元格中第一个超链接要显示的文本*/
function test() {
    let hl = Range("B5").Hyperlinks.Item(1)
    hl.TextToDisplay = "点击跳转"
}
```
