---
title: "Title (属性)"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/AllowEditRange/%E5%B1%9E%E6%80%A7/Title%20%E5%B1%9E%E6%80%A7.html
updated: 1789700916000
hash: CIEtAfFm
---

# Title (属性)

返回或设置受保护的工作表上可编辑的单元格区域的标题。**String** 型，可读写。

## 示例

```javascript
/*本示例将受保护的活动工作表上第二个可编辑的单元格区域的标题设置为“AllowEditRange2”。*/
function test() {
    ActiveSheet.Protection.AllowEditRanges.Item(2).Title = "AllowEditRange2"
}
```

```javascript
/*本示例显示受保护的第一张工作表上第一个可编辑的单元格区域的标题。*/
function test() {
    console.log(Worksheets.Item(1).Protection.AllowEditRanges.Item(1).Title)
}
```
