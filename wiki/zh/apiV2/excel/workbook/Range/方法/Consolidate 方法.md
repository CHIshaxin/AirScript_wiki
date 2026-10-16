---
title: "Consolidate (方法)"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/Range/%E6%96%B9%E6%B3%95/Consolidate%20%E6%96%B9%E6%B3%95.html
updated: 1789700916000
hash: CZ4yRFmQ
---

# Consolidate (方法)

将多个工作表中多个区域的数据合并计算至单个工作表上的单个区域。**Variant** 类型。

## 参数

| 属性 | 数据类型 | 必填 | 说明 |
| --- | --- | --- | --- |
| Sources | any | 可选 | 以文本引用字符串数组的形式给出合并计算的源，该数组采用 R1C1-样式表示法。这些引用必须包含将要合并计算的工作表的完整路径。 |
| Function | any | 可选 | **[XlConsolidationFunction](./../../枚举/XlConsolidationFunction 枚举.html)** 常量之一，用于指定合并计算的类型。 |
| TopRow | any | 可选 | 如果为 **True**，则基于合并计算区域中首行内的列标题对数据进行合并。如果为 **False**，则按位置进行合并计算。默认值为 **False**。 |
| LeftColumn | any | 可选 | 如果为 **True** 则基于合并计算区域中左列内的行标题对数据进行合并计算。如果为 **False**，则按位置进行合并计算。默认值为 **False**。 |
| CreateLinks | any | 可选 | 如果为 **True**，则让合并计算使用工作表链接。如果为 **False**，则让合并计算复制数据。默认值为 **False**。 |

## 返回值

Variant

## 示例

```javascript
/*本示例使用 SUM 函数对工作表 Sheet2 和工作表 Sheet3 的数据进行合并计算，将结果放到工作表 Sheet1 中。*/
function test() {
    Worksheets.Item("Sheet1").Range("A1").Consolidate(["Sheet2!R1C1:R37C6", "Sheet3!R1C1:R37C6"], xlSum)
}
```

```javascript
/*本示例使用 Max 函数求出工作表 Sheet2 和工作表 Sheet3 的数据中的最大值，将结果放到工作表 Sheet1 中。*/
function test() {
    Worksheets.Item("Sheet1").Range("A1").Consolidate(["Sheet2!R1C1:R5C3", "Sheet3!R1C1:R5C3"], xlMax)
}
```
