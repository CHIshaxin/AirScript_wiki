---
title: "Merge (方法)"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/Range/%E6%96%B9%E6%B3%95/Merge%20%E6%96%B9%E6%B3%95.html
updated: 1789700916000
hash: 1XjPl8Tw
---

# Merge (方法)

由指定的 **[Range](./../../Range/Range 对象.html)** 对象创建合并单元格。

## 说明

合并区域的值在该区域左上角的单元格中指定。

## 参数

| 属性 | 数据类型 | 必填 | 说明 |
| --- | --- | --- | --- |
| Across | any | 可选 | 如果为 True，则将指定区域中每一行的单元格合并为一个单独的合并单元格。默认值是 False。 |

## 示例

```javascript
/*本示例将 A1:B3 单元格区域合并*/
function test() {
    Range("A1:B3").Merge()
}
```

```javascript
/*本示例将 A1:B3 单元格区域内单元格逐行合并*/
function test() {
    Range("A1:B3").Merge(true)
}
```
