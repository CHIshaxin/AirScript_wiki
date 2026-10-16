---
title: "Type (属性)"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/Hyperlink/%E5%B1%9E%E6%80%A7/Type%20%E5%B1%9E%E6%80%A7.html
updated: 1789700916000
hash: DKjpP8v4
---

# Type (属性)

返回一个包含 **MsoHyperlinkType** 常量的 **Long** 值，它代表 HTML 框架的位置。

## 示例

```javascript
/*本示例判断工作表Sheet1中第一项超链接是否在单元格中*/
function test() {
    let hl = Application.Worksheets.Item("Sheet1").Hyperlinks.Item(1)
    if (hl.Type == msoHyperlinkRange) {
        console.log('超链接在单元格中')
    } else {
        console.log('超链接不在单元格中')
    }
}
```

```javascript
/*本示例显示第一个形状的超链接是否属于 Shape 对象。*/
function test() {
    let hl = Application.Worksheets.Item(1).Shapes.Item(1).Hyperlink
    console.log(hl.Type == msoHyperlinkShape)
}
```
