---
title: "Insert 方法"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/Range/%E6%96%B9%E6%B3%95/Insert%20%E6%96%B9%E6%B3%95.html
---
# Insert (方法)

在工作表或宏表中插入一个单元格或单元格区域，其他单元格相应移位以腾出空间。

## 参数

| 属性 | 数据类型 | 必填 | 说明 |
| --- | --- | --- | --- |
| Shift | any | 可选 | 指定单元格的调整方式。可为以下 XlInsertShiftDirection 常量之一：xlShiftToRight 或 xlShiftDown。如果省略此参数，ET 将根据区域的形状确定调整方式。 |
| CopyOrigin | any | 可选 | 复制源；即从何处复制插入单元格的格式。 可以是以下 XlInsertFormatOrigin 常量之一： xlFormatFromLeftOrAbove (默认) 或 xlFormatFromRightOrBelow。 |

## 示例

```javascript
/*本示例将 A1 单元格的值设置为 a1，然后在 A1 单元格位置插入一个单元格，原单元格下移*/
function test() {
    Range("A1").Value2 = 'a1'
    Range("A1").Insert(xlShiftDown)
}
```

```javascript
/*本示例将 A1 单元格的值设置为 a1 并加粗，在 A1 单元格位置插入一个单元格，格式参考原A1单元格，原单元格右移*/
function test() {
    Range("A1").Value2 = 'a1'
    Range("A1").Font.Bold = true
    Range("A1").Insert(xlShiftToRight, xlFormatFromRightOrBelow)
}
```
