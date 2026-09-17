---
title: "AddComment 方法"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/Range/%E6%96%B9%E6%B3%95/AddComment%20%E6%96%B9%E6%B3%95.html
---
# AddComment (方法)

为区域添加批注。

## 参数

| 属性 | 数据类型 | 必填 | 说明 |
| --- | --- | --- | --- |
| Text | any | 可选 | 批注文字。 |

## 返回值

Comment

## 示例

```javascript
/*本示例为第一张工作表中 E5 单元格添加批注。*/
function test() {
    Worksheets.Item(1).Range("E5").AddComment("Current Sales")
}
```

```javascript
/*本示例为 A1 单元格添加批注 a1*/
function test() {
    Application.Cells.Item(1, 1).AddComment("a1")
}
```
