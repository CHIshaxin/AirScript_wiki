---
title: "ShowValueFieldButtons (属性)"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/Chart/%E5%B1%9E%E6%80%A7/ShowValueFieldButtons%20%E5%B1%9E%E6%80%A7.html
updated: 1789700916000
hash: BGjvOkgY
---

# ShowValueFieldButtons (属性)

返回或设置是否要在数据透视图上显示值字段按钮。可读/写。

## 说明

将 **ShowValueFieldButtons** 属性设置为 **True** 将在指定的数据透视图上显示**“值字段”**按钮。将该属性设置为 **False** 将隐藏这些按钮。

**ShowValueFieldButtons** 属性对应于**“分析”**选项卡（在选择数据透视图时可用）的**“字段按钮”**下拉列表中的**“显示值字段按钮”**命令。

## 返回值

Boolean

## 示例

```javascript
/*本示例将工作表 Sheet 2 第一个图表工作表设置为显示值字段按钮。*/
function test() {
    let chart = Application.Sheets.Item("Sheet2").ChartObjects(1).Chart
    chart.ShowValueFieldButtons = true
}
```

```javascript
/*本示例显示图表工作表 Chart1 是否显示值字段按钮，并通知用户。*/
function test() {
    let chart = Application.Charts.Item("Chart1").ChartObjects(1).Chart
    if (chart.ShowValueFieldButtons) {
        console.log("图表显示值字段按钮")
    } else {
        console.log("图表不显示值字段按钮")
    }
}
```
