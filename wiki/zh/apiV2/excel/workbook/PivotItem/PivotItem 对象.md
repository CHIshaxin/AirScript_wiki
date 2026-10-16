---
title: "PivotItem (对象)"
source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/PivotItem/PivotItem%20%E5%AF%B9%E8%B1%A1.html
updated: 1789700916000
hash: BcOK46m3
---

# PivotItem (对象)

代表数据透视表字段中的项目。

## 说明

这些项目是某个字段类型中的各个数据项。 **[PivotItem](./../PivotItem/PivotItem 对象.html)** 对象是 **[PivotItems](./../PivotItems/PivotItems 对象.html)** 集合的成员。 **[PivotItems](./../PivotItems/PivotItems 对象.html)** 集合包含某个 **[PivotField](./../PivotField/PivotField 对象.html)** 对象中所有的项目。使用 **PivotItems**(*index*)（其中 *index* 是项目索引号或名称）可返回一个 **[PivotItem](./../PivotItem/PivotItem 对象.html)** 对象。

## 示例

```javascript
/*本示例隐藏“Sheet1”上第一张数据透视表中“name”字段中包含“张一”的所有数据项。*/
function test() {
    Worksheets.Item("Sheet1").PivotTables(1).PivotFields("name").PivotItems("张一").Visible = false
}
```

```javascript
/*本示例显示活动工作表上数据透视表的字段“name”中未展示明细数据的数据项的名称。*/
function test() {
    let pvtItems = ActiveSheet.Range("I1").PivotTable.PivotFields("name").PivotItems()
    for (let i = 1; i <= pvtItems.Count; i++) {
        if (pvtItems.Item(i).ShowDetail == false) {
            console.log(pvtItems.Item(i).Name)
        }
    }
}
```
