---
title: "Add (方法)"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/ChartObjects/%E6%96%B9%E6%B3%95/Add%20%E6%96%B9%E6%B3%95.html
updated: 1789700916000
hash: nMCanCaz
---

# Add (方法)

创建新的嵌入式图表。

## 参数

| 属性 | 数据类型 | 必填 | 说明 |
|---|---|---|---|
| Left | double | 必填 | 以磅为单位给出新对象的初始坐标，该坐标是相对于工作表上单元格 A1 的左上角或图表的左上角的坐标。 |
| Top | double | 必填 | 以磅为单位给出新对象的初始坐标，该坐标是相对于工作表上单元格 A1 的左上角或图表的左上角的坐标。 |
| Width | double | 必填 | 以磅为单位给出新对象的初始大小。 |
| Height | double | 必填 | 以磅为单位给出新对象的初始大小。 |

## 返回值

ChartObject

## 示例

```javascript
/*此示例在工作表 Sheet1 中创建新的嵌入式图表。*/
function test() {
    let chartobject = Application.Sheets.Item("Sheet1").ChartObjects().Add(50, 40, 200, 100)
    chartobject.Chart.ChartWizard(Worksheets.Item("Sheet1").Range("A1:B2"), xlColumn, 6, xlColumns, 1, 0, 1)
}
```

```javascript
/*此示例在第一张工作表中创建新的嵌入式图表，并使用圆角。*/
function test() {
    let chartobject = Application.Sheets.Item(1).ChartObjects().Add(100, 0, 300, 155.5)
    chartobject.RoundedCorners = true
}
```
