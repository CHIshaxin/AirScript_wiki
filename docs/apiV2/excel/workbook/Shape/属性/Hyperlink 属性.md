---
title: "Hyperlink 属性"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/Shape/%E5%B1%9E%E6%80%A7/Hyperlink%20%E5%B1%9E%E6%80%A7.html
---
# Hyperlink (属性)

返回一个 **[Hyperlink](./../../Hyperlink/Hyperlink 对象.html)** 对象，该对象表示该形状的超链接。

## 示例

```javascript
/*本示例加载附于（第一张工作表中）第一个形状的超链接的文档。*/
function test() {
    Worksheets(1).Shapes(1).Hyperlink.Follow(true)
}
```

```javascript
/*本示例显示附于活动工作表第一个形状的超链接地址。*/
function test() {
    let shape = ActiveSheet.Shapes.Item(1)
    console.log(shape.Hyperlink.Address)
}
```
