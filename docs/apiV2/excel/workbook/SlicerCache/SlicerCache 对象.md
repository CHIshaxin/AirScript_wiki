---
title: "SlicerCache 对象"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/SlicerCache/SlicerCache%20%E5%AF%B9%E8%B1%A1.html
---
# SlicerCache (对象)

表示切片器的当前筛选状态，以及有关切片器连接到哪个 **PivotCache** 或 **WorkbookConnection** 的信息。

## 说明

使用 **[Workbook](./../Workbook/Workbook 对象.html)** 对象的 **SlicerCaches** 属性可访问工作簿中 **[SlicerCache](./../SlicerCache/SlicerCache 对象.html)** 对象的集合。

每个切片器都有一个基 **[SlicerCache](./../SlicerCache/SlicerCache 对象.html)** 对象，该对象表示在切片器中显示的项目，以及与对应的项目标题一起显示的平铺的当前用户界面状态。用户在 ET 中看到的每个切片器控件由带有关联的 **[SlicerCache](./../SlicerCache/SlicerCache 对象.html)** 对象的 **[Slicer](./../Slicer/Slicer 对象.html)** 对象表示。

## 示例

```javascript
/*本示例在第一张工作簿上添加新的切片器缓存，并在工作表 Sheet2 中创建切片器。*/
function test() {
    let pvtTable = Worksheets.Item("Sheet1").Range("I1").PivotTable
    let sliCache = Workbooks.Item(1).SlicerCaches.Add2(pvtTable, "name", "Slicer_name")
    sliCache.Slicers.Add(Worksheets.Item("Sheet2"), null, "Name")
}
```

```javascript
/*本示例显示活动工作簿上所有切片器缓存的名称。*/
function test() {
    for (let i = 1; i <= ActiveWorkbook.SlicerCaches.Count; i++) {
        console.log(ActiveWorkbook.SlicerCaches(i).Name)
    }
}
```
