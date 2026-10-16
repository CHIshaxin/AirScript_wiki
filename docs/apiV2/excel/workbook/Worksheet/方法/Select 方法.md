---
title: "Select 方法"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/Worksheet/%E6%96%B9%E6%B3%95/Select%20%E6%96%B9%E6%B3%95.html
---
# Select (方法)

选择对象。

## 说明

要选择单元格或单元格区域，请使用 **Select** 方法。要使单个单元格成为活动单元格，请使用**Activate**方法。

## 参数

| 属性 | 数据类型 | 必填 | 说明 |
| --- | --- | --- | --- |
| Replace | any | 可选 | （仅用于工作表）。如果为 True，则用指定的对象替换当前所选内容。如果为 False，则扩展当前所选内容以包括以前选择的对象和指定的对象。 |

## 示例

```javascript
/*本示例选定第3张工作表。*/
function test() {
    Worksheets.Item(3).Select()
}
```

```javascript
/*本示例扩展当前所选的工作表，增加选择 Sheet1 工作表。*/
function test() {
    Worksheets.Item("Sheet1").Select(false)
}
```
