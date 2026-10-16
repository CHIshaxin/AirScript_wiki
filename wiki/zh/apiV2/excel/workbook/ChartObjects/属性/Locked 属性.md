---
title: "Locked (属性)"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/ChartObjects/%E5%B1%9E%E6%80%A7/Locked%20%E5%B1%9E%E6%80%A7.html
updated: 1789700916000
hash: -nUzrRvp
---

# Locked (属性)

返回或设置一个 **Boolean** 值，它指明对象是否已被锁定。

## 说明

如果对象已被锁定，此属性将返回 **True**；如果在工作表处于受保护状态时仍能修改对象，则返回 **False**。

## 示例

```javascript
/*此示例取消对工作表 Sheet1 中所有图表的锁定，并通知用户。*/
function test() {
    let chartobjects = Application.Worksheets.Item("Sheet1").ChartObjects()
    chartobjects.Locked = false
    console.log("已取消对工作表 Sheet1 上所有图表的锁定")
}
```

```javascript
/*此示例显示工作表 Sheet2 中所有图表是否已被锁定。*/
function test() {
    let chartobjects = Application.Sheets.Item("Sheet2").ChartObjects()
    console.log(chartobjects.Locked)
}
```
