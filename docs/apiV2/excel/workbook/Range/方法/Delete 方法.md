---
title: "Delete 方法"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/Range/%E6%96%B9%E6%B3%95/Delete%20%E6%96%B9%E6%B3%95.html
---
# Delete (方法)

删除对象。

## 参数

| 属性 | 数据类型 | 必填 | 说明 |
| --- | --- | --- | --- |
| Shift | any | 可选 | 仅用于 Range 对象。指定如何调整单元格以替换删除的单元格。可以为以下 XlDeleteShiftDirection 常量之一：xlShiftToLeft 或 xlShiftUp。如果省略此参数，ET 将根据区域的形状确定调整方式。 |

## 示例

```javascript
/*本示例设置 A1 单元格值为 a1，A2 单元格值为 a2，将 A1 单元格删除，下方单元格向上移动。*/
function test() {
    Range("A1").Value2 = 'a1'
    Range("A2").Value2 = 'a2'
    Range("A1").Delete(xlShiftUp)
}
```

```javascript
/*本示例设置 A1 单元格值设置为 a1，B1 单元格的值设为 b1，并将 A 列删除*/
function test() {
    Range("A1").Value2 = "a1"
    Range("B1").Value2 = "b1"
    Range("A:A").Delete()
}
```
