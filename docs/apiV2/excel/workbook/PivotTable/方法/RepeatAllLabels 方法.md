---
title: "RepeatAllLabels 方法"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/PivotTable/%E6%96%B9%E6%B3%95/RepeatAllLabels%20%E6%96%B9%E6%B3%95.html
---
# RepeatAllLabels (方法)

指定是否要对指定数据透视表中的所有透视字段重复项目标签。

## 说明

使用 **RepeatAllLabels** 方法对应于 **“数据透视表工具设计”** 选项卡的 **“报表布局”** 下拉列表中的 **“重复所有项目标签”** 和 **“不重复项目标签”** 命令。

若要指定是否要对单个数据透视表重复项目标签，请使用 **RepeatLabels** 属性。

## 参数

| 属性 | 数据类型 | 必填 | 说明 |
| --- | --- | --- | --- |
| Repeat | XlPivotFieldRepeatLabels | 必填 | 指定是否重复数据透视表中的所有字段项目标签。 |

## 示例

```javascript
/*本示例将第一张工作表第一张数据透视表设置为不重复数据透视表中的所有字段项目标签。*/
function test() {
    Worksheets.Item(1).PivotTables(1).RepeatAllLabels(xlDoNotRepeatLabels)
}
```

```javascript
/*本示例将活动工作表第一张数据透视表设置为重复数据透视表中的所有字段项目标签。*/
function test() {
    ActiveSheet.PivotTables(1).RepeatAllLabels(xlRepeatLabels)
}
```
