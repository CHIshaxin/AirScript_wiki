---
title: "AddTop10 (方法)"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/FormatConditions/%E6%96%B9%E6%B3%95/AddTop10%20%E6%96%B9%E6%B3%95.html
updated: 1789700916000
hash: D6lCwkA5
---

# AddTop10 (方法)

返回一个代表指定区域条件格式规则的 **[Top10](./../../Top10/Top10 对象.html)** 对象。

## 说明

使用 **[Top10](./../../Top10/Top10 对象.html)** 对象，可以基于指定的截止值查找单元格区域中的最高值和最低值。例如，可以查找区域报告中位居前五位的销售产品、客户调查中位居最后百分之十五的产品，或者部门人员分析中位居前 25 位的薪金。

## 示例

```javascript
/*本示例在活动工作表上单元格区域 F1:F10 新增Top10条件格式。*/
function test() {
    let top = ActiveSheet.Range("F1:F10").FormatConditions.AddTop10()
    top.TopBottom = xlTop10Top
    top.Rank = 5
    top.Percent = false
    top.Font.Bold = true
}
```

```javascript
/*本示例在活动工作表上单元格区域 A1:A10 新增Top10条件格式。*/
function test() {
    let top = ActiveSheet.Range("A1:A10").FormatConditions.AddTop10()
    top.TopBottom = xlTop10Bottom
    top.Rank = 3
    top.Percent = false
    top.Interior.ColorIndex = 5
}
```
