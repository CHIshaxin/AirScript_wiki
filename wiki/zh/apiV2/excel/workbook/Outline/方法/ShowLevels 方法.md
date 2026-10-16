---
title: "ShowLevels (方法)"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/Outline/%E6%96%B9%E6%B3%95/ShowLevels%20%E6%96%B9%E6%B3%95.html
updated: 1789700916000
hash: DSXJsIEJ
---

# ShowLevels (方法)

显示指定行号和/或列号在分级显示中的层次。

## 说明

必须至少指定一个参数。

## 参数

| 属性 | 数据类型 | 必填 | 说明 |
| --- | --- | --- | --- |
| RowLevels | any | 可选 | 指定分级显示中的行层次。如果该分级显示包含的层次数少于指定层次，则 ET 显示所有层次。如果该参数为 0（零）或者省略该参数，则不对行采取任何操作。 |
| ColumnLevels | any | 可选 | 指定分级显示中的列层次。如果该分级显示包含的层次数少于指定层次，则 ET 显示所有层次。如果该参数为 0（零）或者省略该参数，则不对列采取任何操作。 |

## 返回值

Variant

## 示例

```javascript
/*本示例对工作表 Sheet1 的第一行到第三行的行层次和第一列的列层次进行分级显示。*/
function test() {
    Worksheets.Item("Sheet1").Outline.ShowLevels(3, 1)
}
```

```javascript
/*本示例对活动工作表的第一行的行层次和第一列的列层次进行分级显示。*/
function test() {
    ActiveSheet.Outline.ShowLevels(1, 1)
}
```
