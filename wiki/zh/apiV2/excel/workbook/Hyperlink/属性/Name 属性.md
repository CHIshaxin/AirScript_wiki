---
title: "Name (属性)"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/Hyperlink/%E5%B1%9E%E6%80%A7/Name%20%E5%B1%9E%E6%80%A7.html
updated: 1789700916000
hash: BDAxcBRc
---

# Name (属性)

返回一个 **String** 值，它代表对象的名称。

## 示例

```javascript
/*本示例显示工作表Sheet1中第一项超链接的名称*/
function test() {
    let hl = Application.Worksheets.Item("Sheet1").Hyperlinks.Item(1)
    console.log(`超链接的名称是 ${hl.Name}`)
}
```

```javascript
/*本示例显示单元格B2中第一项超链接的名称*/
function test() {
    let hl = Range("B2").Hyperlinks.Item(1)
    console.log(`超链接的名称是 ${hl.Name}`)
}
```
