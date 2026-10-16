---
title: "RemoveDuplicates (方法)"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/Range/%E6%96%B9%E6%B3%95/RemoveDuplicates%20%E6%96%B9%E6%B3%95.html
updated: 1789700916000
hash: UI-CQzlg
---

# RemoveDuplicates (方法)

从值区域中删除重复的值。

## 参数

| 属性 | 数据类型 | 必填 | 说明 |
| --- | --- | --- | --- |
| Columns | any | 可选 | 包含重复信息的列的索引数组。 |
| Header | Default | 可选 | 指定第一行是否包含标题信息。**xlNo** 是默认值；如果希望 ET 确定标题，则指定 **xlGuess**。 |

## 示例

```javascript
/*以下代码示例从区域所有列中删除重复值。*/
function test() {
    ActiveSheet.Range("A1:C100").RemoveDuplicates([1])
}
```

```javascript
/*以下代码示例删除前两列的重复值（指定第一行为标题）*/
function test() {
    Range("A1:B100").RemoveDuplicates([1, 2], xlYes)
}
```
