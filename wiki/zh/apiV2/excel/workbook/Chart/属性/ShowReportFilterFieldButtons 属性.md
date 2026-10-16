---
title: "ShowReportFilterFieldButtons (属性)"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/Chart/%E5%B1%9E%E6%80%A7/ShowReportFilterFieldButtons%20%E5%B1%9E%E6%80%A7.html
updated: 1789700916000
hash: DX_tCb1B
---

# ShowReportFilterFieldButtons (属性)

返回或设置是否要在数据透视图上显示报表筛选字段按钮。可读/写。

## 说明

将 **ShowReportFilterFieldButtons** 属性设置为 **True** 将在指定的数据透视图上显示**“报表筛选字段”**按钮。将该属性设置为 **False** 将隐藏这些按钮。

**ShowReportFilterFieldButtons** 属性对应于**“分析”**选项卡（在选择数据透视图时可用）的**“字段按钮”**下拉列表中的**“显示报表筛选字段按钮”**命令。

## 返回值

Boolean

## 示例

```javascript
/*本示例将图表工作表 Chart 1 设置为显示报表筛选字段按钮。*/
function test() {
    let chart = Application.Charts.Item("Chart1").ChartObjects(1).Chart
    chart.ShowReportFilterFieldButtons = true
}
```

```javascript
/*本示例判断工作表 Sheet2 中第一个内嵌图表是否显示报表筛选字段按钮，并通知用户。*/
function test() {
    let chart = Application.Sheets.Item("Sheet2").ChartObjects(1).Chart
    if (chart.ShowReportFilterFieldButtons) {
        console.log("图表显示报表筛选字段按钮")
    } else {
        console.log("图表不显示报表筛选字段按钮")
    }
}
```
