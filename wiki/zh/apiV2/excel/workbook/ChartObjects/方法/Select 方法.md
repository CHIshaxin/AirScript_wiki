---
title: "Select (方法)"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/ChartObjects/%E6%96%B9%E6%B3%95/Select%20%E6%96%B9%E6%B3%95.html
updated: 1789700916000
hash: AT-dPHTI
---

# Select (方法)

选择对象。

## 参数

| 属性 | 数据类型 | 必填 | 说明 |
|---|---|---|---|
| Replace | any | 可选 | 如果为 True，则用指定的对象替换当前所选内容。如果为 False，则扩展当前所选内容以包括以前选择的对象和指定的对象。 |

## 示例

```javascript
/*此示例选择工作表 Sheet1 的第一个图表，再扩展选择 Sheet1 中所有的图表。*/
function test() {
    Application.Worksheets.Item("Sheet1").ChartObjects(1).Select()
    let chartobjects = Application.Worksheets.Item("Sheet1").ChartObjects()
    chartobjects.Select(false)
}
```

```javascript
/*此示例选择第一张工作表中所有图表。*/
function test() {
    let chartobjects = Application.Worksheets.Item(1).ChartObjects()
    chartobjects.Select()
}
```
