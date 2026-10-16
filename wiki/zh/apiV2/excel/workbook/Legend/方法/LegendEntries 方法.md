---
title: "LegendEntries (方法)"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/Legend/%E6%96%B9%E6%B3%95/LegendEntries%20%E6%96%B9%E6%B3%95.html
updated: 1789700916000
hash: DH_N2urK
---

# LegendEntries (方法)

返回表示图例中的单个图例项（ **[LegendEntry](./../../LegendEntry/LegendEntry 对象.html)** 对象）或图例项集合（ **[LegendEntries](./../../LegendEntries/LegendEntries 对象.html)** 对象）的对象。

## 参数

| 属性 | 数据类型 | 必填 | 说明 |
| --- | --- | --- | --- |
| Index | any | 可选 | 图例输入项的编号。 |

## 示例

```javascript
/*本示例设置 Chart1 上第一个图例项的字体。*/
function test() {
    Application.Charts.Item("Chart1").ChartObjects(1).Chart.Legend.LegendEntries(1).Font.Name = "Arial"
}
```

```javascript
/*本示例显示第一张工作表上的第一个图表中图例的图例项数量。*/
function test() {
    console.log(Worksheets.Item(1).ChartObjects(1).Chart.Legend.LegendEntries().Count)
}
```
