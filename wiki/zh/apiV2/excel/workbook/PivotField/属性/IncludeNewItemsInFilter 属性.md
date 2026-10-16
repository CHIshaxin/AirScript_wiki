---
title: "IncludeNewItemsInFilter (属性)"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/PivotField/%E5%B1%9E%E6%80%A7/IncludeNewItemsInFilter%20%E5%B1%9E%E6%80%A7.html
updated: 1789700916000
hash: OJWMR0mX
---

# IncludeNewItemsInFilter (属性)

在将手动筛选应用于透视字段时，此属性允许开发人员指定是应跟踪排除的项目还是应跟踪包含的项目。**Boolean** 类型，可读写。

## 说明

此属性的默认值为 **False**。

应用手动筛选后，开发人员可以将 **IncludeNewItemsInFilter** 属性设置为 **True** 来跟踪排除的项目，也可以将此属性设置为 **False** 来跟踪包含的项目。如果在 **IncludeNewItemsInFilter** 属性设置为 **True** 后向源数据中添加了新项目，则这些新项目只有在进行新的刷新操作后才会显示在数据透视表中，因为它们未处在 ET 正在跟踪的项目集合中。

对于 OLAP 层次结构，此设置在 **CubeField** 对象上设置，而试图在作为层次结构一部分的透视字段上设置它将会失败（并会产生运行时错误）。可以通过作为层次结构一部分的透视字段获取此设置，而且它将返回与对应的 **CubeField.IncludeNewItemsInFilter** 属性相同的结果。

切换此设置后，以下集合将为空：**HiddenItemsList**、**HiddenItems**、**VisibleItemsList** 和 **VisibleItems**。 当 **IncludeNewItemsInFilter** 设置为 **False** 时，**HiddenItemsList** 和 **HiddenItems** 集合为空，并且无法向它们添加项目。试图添加项目时会返回运行时错误。 当 **IncludeNewItemsInFilter** 设置为 **True** 时，**VisibleItemsList** 和 **VisibleItems** 集合为空，并且无法向它们添加项目。试图添加项目时会返回运行时错误。

在数据透视表中，此设置在 **[PivotField](./../../PivotField/PivotField 对象.html)** 对象上设置。切换此设置将不会改变筛选器状态。

## 示例

```javascript
/*本示例判断工作表 Sheet1 中数据透视表的字段“score”应用手动筛选时，是应跟踪排除的项目还是跟踪包含的项目，并通知用户。*/
function test() {
    let pvtField = Worksheets.Item("Sheet1").Range("I1").PivotTable.PivotFields("score")
    if (pvtField.IncludeNewItemsInFilter) {
        console.log("跟踪排除的项目")
    } else {
        console.log("跟踪包含的项目")
    }
}
```

```javascript
/*本示例使活动工作表中第一张数据透视表的字段“name”应用手动筛选时，指定跟踪排除的项目。*/
function test() {
    ActiveSheet.PivotTables(1).PivotFields("name").IncludeNewItemsInFilter = true
}
```
