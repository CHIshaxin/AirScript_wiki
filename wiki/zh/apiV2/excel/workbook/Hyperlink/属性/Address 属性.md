---
title: "Address (属性)"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/Hyperlink/%E5%B1%9E%E6%80%A7/Address%20%E5%B1%9E%E6%80%A7.html
updated: 1789700916000
hash: DRgn4W9h
---

# Address (属性)

返回或设置一个 **String** 型，它代表目标文档的地址。

## 示例

```javascript
/*本示例显示单元格B2中第一项超链接的地址*/
function test() {
    let hl = Range("B2").Hyperlinks.Item(1)
    console.log(`超链接的目标地址是 ${hl.Address}`)
}
```

```javascript
/*本示例修改工作表Sheet1中第一项超链接的地址*/
function test() {
    let hl = Application.Worksheets.Item("Sheet1").Hyperlinks.Item(1)
    hl.Address = "https://wps.cn/"
    console.log(`超链接的地址已修改为 ${hl.Address}`)
}
```
