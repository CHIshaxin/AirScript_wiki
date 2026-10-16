---
title: "ShowDetail (属性)"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/Range/%E5%B1%9E%E6%80%A7/ShowDetail%20%E5%B1%9E%E6%80%A7.html
updated: 1789700916000
hash: KPnbWR4A
---

# ShowDetail (属性)

如果扩展了指定区域的分级显示（从而行或列的明细数据可见），则为 **True**。指定区域必须为分级显示中的单个汇总列或汇总行。**Variant** 型，可读写。对于 **[PivotItem](./../../PivotItem/PivotItem 对象.html)** 对象（如果该区域在数据透视表中，则为 **[Range](./../../Range/Range 对象.html)** 对象），当数据项显示明细数据时，此属性设为 **True**。

## 说明

此属性不可用于 OLAP （OLAP：为查询和报表（而不是处理事务）而进行了优化的数据库技术。OLAP 数据是按分级结构组织的，它存储在多维数据集而不是表中。）数据源。

如果指定区域不在数据透视表中，则下列声明为真：

- 指定区域必须在单个汇总行或汇总列中。
- 如果指定行或列的任意子级处于隐藏状态，则此属性返回 **False**。
- 将此属性设置为 **True** 相当于显示指定汇总行或汇总列的所有子级内容。
- 将此属性设置为 **False** 相当于隐藏汇总行或汇总列的所有子级内容。

如果指定区域为数据透视表，则当该区域连续时，就可以同时对多个单元格设置此属性。仅当指定区域为单个单元格时，才能返回此属性的值。

## 示例

```javascript
/*此示例显示工作表 Sheet1 上分级显示的汇总行的明细数据。运行此示例之前，应先创建包含单个汇总行的简单分级显示，然后折叠该分级显示，使之仅显示汇总行。选定汇总行中的某一单元格，再运行此示例。*/
function test() {
    Worksheets.Item("Sheet1").Activate()
    let myRange = ActiveCell.CurrentRegion
    let lastRow = myRange.Rows.Count
    myRange.Rows.Item(lastRow).ShowDetail = true
}
```

```javascript
/*本示例选中 A40 单元格，判断选中汇总行的明细数据是否显示，如果显示，则隐藏；如果影藏，则显示*/
function test() {
    Range("A40").Select()
    let rows = ActiveCell.Rows
    if (rows.ShowDetail) {
        rows.ShowDetail = false
        console.log("明细数据已隐藏")
    } else {
        rows.ShowDetail = true
        console.log("明细数据已显示")
    }
}
```
