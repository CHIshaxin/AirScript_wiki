---
title: "Has3DEffect (属性)"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/Series/%E5%B1%9E%E6%80%A7/Has3DEffect%20%E5%B1%9E%E6%80%A7.html
updated: 1789700916000
hash: BL1_vR2w
---

# Has3DEffect (属性)

如果数据系列有三维外观，则为 **True**。可读/写 **Boolean** 类型。

## 说明

该属性仅适用于气泡图。

## 示例

```javascript
/*此示例显示活动工作表中第一个图表的第二个数据系列是否有三维外观。*/
function test() {
    console.log(ActiveSheet.ChartObjects(1).Chart.SeriesCollection(2).Has3DEffect)
}
```

```javascript
/*此示例显示图表工作表 Chart1 上的第三个数据系列是否有三维外观。*/
function test() {
    let series3 = Application.Charts.Item("Chart1").ChartObjects(1).Chart.SeriesCollection(3)
    console.log(series3.Has3DEffect)
}
```
