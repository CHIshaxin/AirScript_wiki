---
title: "Names (属性)"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/Workbook/%E5%B1%9E%E6%80%A7/Names%20%E5%B1%9E%E6%80%A7.html
updated: 1789700916000
hash: CmXW9dnZ
---

# Names (属性)

返回一个 **[Names](./../../Names/Names 对象.html)** 集合，它代表指定工作簿的所有名称（包括所有指定工作表的名称）。 **[Names](./../../Names/Names 对象.html)** 对象，只读。

## 说明

在不使用对象识别符的情况下使用此属性等效于使用 `ActiveWorkbook.Names`。

## 示例

```javascript
/*本示例是将 Sheet1 中的 A1 单元格的名称定义为“myName”。*/
function test() {
    Application.ActiveWorkbook.Names.Add("myName", null, null, null, null, null, null, null, null, "=Sheet1!R1C1")
}
```

```javascript
/*本示例是将 Sheet1 中的 C1 单元格的名称定义为“myCell”。*/
function test() {
    Application.ActiveWorkbook.Names.Add("myCell", "=Sheet1!$C$1")
}
```
