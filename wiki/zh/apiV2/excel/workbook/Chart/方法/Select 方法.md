---
title: "Select (方法)"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/Chart/%E6%96%B9%E6%B3%95/Select%20%E6%96%B9%E6%B3%95.html
updated: 1789700916000
hash: CJ4cA7j5
---

# Select (方法)

选择对象。

## 说明

要选择图表，请使用 **Select** 方法。要使单个图表成为活动图表，请使用 **Activate** 方法。

## 参数

| 属性 | 数据类型 | 必填 | 说明 |
|---|---|---|---|
| Replace | any | 可选 | （仅用于工作表）。如果为 True，则用指定的对象替换当前所选内容。如果为 False，则扩展当前所选内容以包括以前选择的对象和指定的对象。 |

## 示例

```javascript
/*本示例选定图表工作表 Chart1。*/
function test() {
    let chart = Application.Charts.Item("Chart1")
    chart.Select()
}
```

```javascript
/*本示例扩展所选的工作表 Sheet1，增加选择 Chart1 工作表。*/
function test() {
    Application.Sheets.Item("Sheet1").Select()
    let chart = Application.Charts.Item("Chart1")
    chart.Select(false)
}
```
