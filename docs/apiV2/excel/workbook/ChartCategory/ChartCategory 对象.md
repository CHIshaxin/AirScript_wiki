---
title: "ChartCategory 对象"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/ChartCategory/ChartCategory%20%E5%AF%B9%E8%B1%A1.html
---
# ChartCategory (对象)

指定图表类型的类别。

## 示例

```javascript
/*此示例判断工作表 Sheet1 第一个内嵌图表第一个图表组第一个类别是否筛选出序列，并通知用户。*/
function test() {
    let chart = Application.Worksheets.Item("Sheet1").ChartObjects(1).Chart.ChartGroups(1)
    if (chart.CategoryCollection(1).IsFiltered) {
        console.log("图表组的第一个类别筛选出序列")
    } else {
        console.log("图表组的第一个类别未筛选出序列")
    }
}
```

```javascript
/*此示例显示工作表 Sheet1 第一个内嵌图表第一个图表组的第四个类别的实际类型。*/
function test() {
    let chart = Application.Worksheets.Item("Sheet1").ChartObjects(1).Chart.ChartGroups(1)
    let category = chart.FullCategoryCollection(4)
    console.log(category.constructor.name)
}
```
