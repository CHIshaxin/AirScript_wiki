---
title: "Name (属性)"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/Chart/%E5%B1%9E%E6%80%A7/Name%20%E5%B1%9E%E6%80%A7.html
updated: 1789700916000
hash: CRKE5T1j
---

# Name (属性)

返回或设置一个 **String** 值，它代表对象的名称。

## 说明

此属性对图表对象（嵌入式图表）而言是只读的。

## 示例

```javascript
/*本示例通过名称Chart1获取图表工作表，并显示名称，然后将该图表名称修改为成绩图表。*/
function test() {
    let chart = Application.Charts.Item("Chart1")
    console.log(chart.Name)
    chart.Name = "成绩图表"
}
```

```javascript
/*本示例显示活动工作簿上所有工作表上第一张内嵌图表的名称。*/
function test() {
    for (let i = 1; i <= Sheets.Count; i++) {
        let sheet = Application.Sheets.Item(i)
        if (sheet.ChartObjects().Count >= 1) {
            let chart = sheet.ChartObjects(1).Chart
            console.log(chart.Name)
        }
    }
}
```
