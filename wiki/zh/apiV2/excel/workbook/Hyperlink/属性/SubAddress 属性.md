---
title: "SubAddress (属性)"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/Hyperlink/%E5%B1%9E%E6%80%A7/SubAddress%20%E5%B1%9E%E6%80%A7.html
updated: 1789700916000
hash: C106slY7
---

# SubAddress (属性)

返回或设置文档中与指定超链接相关联的位置。**String** 类型，可读写。

## 示例

```javascript
/*本示例将一个区域位置添加到第一个形状的超链接中*/
function test() {
    Application.Worksheets.Item(1).Shapes.Item(1).Hyperlink.SubAddress = "A1:B10"
}
```

```javascript
/*本示例将单元格G12的第一个超链接的子地址设置为"/home/"。*/
function test() {
    let hl = Range("G12").Hyperlinks.Item(1)
    hl.SubAddress = "/home/"
}
```
