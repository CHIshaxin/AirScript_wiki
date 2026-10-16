---
title: "ScreenTip (属性)"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/Hyperlink/%E5%B1%9E%E6%80%A7/ScreenTip%20%E5%B1%9E%E6%80%A7.html
updated: 1789700916000
hash: F3a8G06B
---

# ScreenTip (属性)

返回或设置指定超链接的屏幕提示文字。**String** 类型，可读写。

## 说明

在将文档保存到网页中之后，如果在 Web 浏览器中查看该文档，则当鼠标移动到超链接上时，“屏幕提示”文字就会显示。某些浏览器可能不支持“屏幕提示”。

## 示例

```javascript
/*本示例修改活动工作表上第一个超链接的屏幕提示*/
function test() {
    Application.ActiveSheet.Hyperlinks.Item(1).ScreenTip = "Return to the home page"
}
```

```javascript
/*本示例修改C7单元格的第一个超链接的屏幕提示并显示出来*/
function test() {
    let hl = Range("C7").Hyperlinks.Item(1)
    hl.ScreenTip = "BACKHOME"
    console.log(`屏幕提示为 ${hl.ScreenTip}`)
}
```
