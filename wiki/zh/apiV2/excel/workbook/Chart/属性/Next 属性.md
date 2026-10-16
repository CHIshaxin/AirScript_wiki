---
title: "Next (属性)"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/Chart/%E5%B1%9E%E6%80%A7/Next%20%E5%B1%9E%E6%80%A7.html
updated: 1789700916000
hash: BK6XuPLE
---

# Next (属性)

返回代表下一个工作表的 **[Worksheet](./../../Worksheet/Worksheet 对象.html)** 对象。

## 说明

如果该对象为区域，则此属性会模拟 Tab 键，但会返回下一个单元格，而不选中它。

在处于保护状态的工作表中，此属性返回下一个未锁定单元格。在未保护的工作表中，此属性总是返回紧靠指定单元格右边的单元格。

## 示例

```javascript
/*本示例显示图表工作表 Chart1 下一个工作表的索引号。*/
function test() {
    let chart = Application.Charts.Item("Chart1")
    console.log(chart.Next.Index)
}
```

```javascript
/*本示例激活图表工作表 Chart1 下一个工作表。*/
function test() {
    let chart = Application.Charts.Item("Chart1")
    chart.Next.Activate()
}
```
