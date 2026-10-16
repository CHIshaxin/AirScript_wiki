---
title: "ClearHyperlinks 方法"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/Range/%E6%96%B9%E6%B3%95/ClearHyperlinks%20%E6%96%B9%E6%B3%95.html
---
# ClearHyperlinks (方法)

删除指定区域中的所有超链接。

## 说明

对指定区域调用 **ClearHyperlinks** 方法相当于在 **“开始”** 选项卡的 **“编辑”** 部分中，使用 **“清除”** 下拉列表中的 **“清除超链接”** 命令。只有超链接才会被删除；所有其他单元格内容（如文本和格式设置）将不受影响。

## 返回值

null

## 示例

```javascript
/*本示例选中 A1 单元格，为选中区域添加超链接 https://www.wps.cn/，再清除超链接*/
function test() {
    Range("A1").Select()
    Selection.Hyperlinks.Add(Selection, "https://www.wps.cn/")
    Selection.ClearHyperlinks()
}
```

```javascript
/*本示例将 B3 单元格设置超链接，并删除 A1:B5 单元格区域的超链接*/
function test() {
    Range("B3").Hyperlinks.Add(Range("B3"), "https://www.wps.com/")
    Range("A1:B5").Rows.ClearHyperlinks()
}
```
