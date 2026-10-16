---
title: "AutoText (属性)"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/DataLabels/%E5%B1%9E%E6%80%A7/AutoText%20%E5%B1%9E%E6%80%A7.html
updated: 1789700916000
hash: RH7ELCG8
---

# AutoText (属性)

如果对象会根据内容自动生成合适的文字，则为 **True**。可读/写 **Boolean** 类型。

## 说明

如果在“即时窗口”中运行 **ActiveChart.SeriesCollection(1).DataLabels().AutoText**，您将收到以下结果：只有在所有 **DataLabels** 的 **AutoText** 都为 **True** 时返回 **True**，如果所有 **DataLabels** 的 **AutoText** 都为 **False**，或者部分 **DataLabels** 的 **AutoText** 为 **False**，那么返回 **False**。

## 示例

```javascript
/*此示例对 Chart1 的第一个数据序列的数据标签进行设置，自动生成合适的文字。*/
function test() {
    let datalabels = Application.Charts.Item("Chart1").ChartObjects(1).Chart.SeriesCollection(1).DataLabels()
    datalabels.AutoText = true
}
```

```javascript
/*此示例显示工作表 Sheet1 中第一张图表的第二个数据序列的数据标签是否自动生成合适的文字。*/
function test() {
    let datalabels = Application.Sheets.Item("Sheet1").ChartObjects(1).Chart.SeriesCollection(2).DataLabels()
    console.log(datalabels.AutoText)
}
```
