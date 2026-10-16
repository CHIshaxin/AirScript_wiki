---
title: "Cut (方法)"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/Range/%E6%96%B9%E6%B3%95/Cut%20%E6%96%B9%E6%B3%95.html
updated: 1789700916000
hash: Cvc9u40P
---

# Cut (方法)

将对象剪切到剪贴板，或者将其粘贴到指定的目的地。

## 说明

剪切的区域必须由相邻的单元格组成。

## 参数

| 属性 | 数据类型 | 必填 | 说明 |
| --- | --- | --- | --- |
| Destination | any | 可选 | 应在其中粘贴对象的目标区域。如果省略此参数，区域对象会被剪切到剪贴板。 |

## 示例

```javascript
/*本示例将 A1 单元格的值剪切到 B1 单元格中*/
function test() {
    Range("A1").Cut(Range("B1"))
}
```

```javascript
/*本示例将 A1 单元格的文本设置为加粗，并剪切到 C1 单元格中*/
function test() {
    Range("A1").Font.Bold = true
    Range("A1").Select()
    Selection.Cut(Range("C1"))
}
```
