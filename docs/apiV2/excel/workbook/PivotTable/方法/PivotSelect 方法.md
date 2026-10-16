---
title: "PivotSelect 方法"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/PivotTable/%E6%96%B9%E6%B3%95/PivotSelect%20%E6%96%B9%E6%B3%95.html
---
# PivotSelect (方法)

选定数据透视表的一部分。

## 说明

只能使用指定模式来选择数据透视表中的相应项。例如，无法通过 **xlButton** 模式来选择数据和标志；同样，也无法通过 **xlDataOnly** 模式来选择按钮。

## 参数

| 属性 | 数据类型 | 必填 | 说明 |
| --- | --- | --- | --- |
| Name | string | 必填 | 要选择的数据透视表的一部分。 |
| Mode | Default<XlPTSelectionMode, 0> | 可选 | 指定结构化选择模式。 |
| UseStandardName | any | 可选 | 如果为 True，则表示将在其他位置运行录制的宏。 |

## 示例

```javascript
/*本示例选定第一张工作表上的第一张数据透视表中的所有日期标志。*/
function test() {
    Worksheets.Item(1).PivotTables(1).PivotSelect("date[All]", xlLabelOnly)
}
```

```javascript
/*本示例选中数据透视表中 y 字段的值*/
function test() {
    ActiveSheet.PivotTables(1).PivotSelect("y", xlDataOnly)
}
```
