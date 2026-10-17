---
title: "Columns (属性)"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/Worksheet/%E5%B1%9E%E6%80%A7/Columns%20%E5%B1%9E%E6%80%A7.html
updated: 1789700916000
hash: BElpv2Un
---

# Columns (属性)

返回一个 Range对象，它代表活动工作表中的所有列。如果活动文档不是工作表，则 **Columns** 属性失效。

## 说明

在不使用对象识别符的情况下使用此属性等效于使用 `ActiveSheet.Columns`。

此属性在应用于一个是多重选定区域的 **[Range](./../../Range/Range 对象.html)** 对象时，会只从该区域的第一个子区域中返回列。例如，如果 **[Range](./../../Range/Range 对象.html)** 对象有两个子区域 A1:B2 和 C3:D4，那么，`Selection.Columns.Count` 的返回值是 2，而不是 4。若要对一个可能包含多重选定区域的区域使用此属性，请测试 `Areas.Count` 以确定此区域内是否包含多个子区域。如果包含，请对此区域内的每个子区域进行循环。

## 示例

```javascript
/*此示例将 Sheet1 的第一列（A 列）的字体设置为加粗。*/
function test() {
    Application.Worksheets.Item("Sheet1").Columns.Item(1).Font.Bold = true
}
```

```javascript
/*本示例将活动工作表的第二列字体设置为8磅。*/
function test() {
    Application.ActiveSheet.Columns.Item(2).Font.Size = 8
}
```
