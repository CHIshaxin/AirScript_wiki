---
title: "RowDifferences (方法)"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/Range/%E6%96%B9%E6%B3%95/RowDifferences%20%E6%96%B9%E6%B3%95.html
updated: 1789700916000
hash: De_F_CKk
---

# RowDifferences (方法)

返回一个 **[Range](./../../Range/Range 对象.html)** 对象，该对象表示每行中与比较单元格内容不同的所有单元格。

## 参数

| 属性 | 数据类型 | 必填 | 说明 |
| --- | --- | --- | --- |
| Comparison | any | 必填 | 用于与指定区域进行比较的单个单元格。 |

## 返回值

Range

## 示例

```javascript
/*本示例选定工作表 Sheet1 上第一行中与单元格 D1 内容不同的单元格。*/
function test() {
    Worksheets.Item("Sheet1").Activate()
    let c1 = ActiveSheet.Rows.Item(1).RowDifferences(ActiveSheet.Range("D1"))
    c1.Select()
}
```

```javascript
/*本示例将 A2:E2 单元格区域的值设置为 2，G2 单元格的值设置为 5，显示 G2 单元格所在行与其值不相等的单元格地址*/
function test() {
    Range("A2:E2").Value2 = 2
    Range("G2").Value2 = 5
    console.log(Rows.Item(Range("G2").Row).RowDifferences(Range("G2")).Address())
}
```
