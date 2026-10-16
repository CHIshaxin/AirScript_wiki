---
title: "RemoveSubtotal 方法"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/Range/%E6%96%B9%E6%B3%95/RemoveSubtotal%20%E6%96%B9%E6%B3%95.html
---
# RemoveSubtotal (方法)

删除列表中的分类汇总。

## 返回值

Variant

## 示例

```javascript
/*本示例删除工作表 Sheet1 上单元格区域 A1:G37 的分类汇总。本示例应在包含分类汇总的数据清单上运行。*/
function test() {
    Worksheets.Item("Sheet1").Range("A1:G37").RemoveSubtotal()
}
```

```javascript
/*本示例删除 A 至 B 列的分类汇总*/
function test() {
    Range("A:B").RemoveSubtotal()
}
```
